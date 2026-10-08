"""Face-triggered G1 presentation controller. Python 3.10+ on Windows or Linux."""
import argparse, json, os, time, urllib.request
from pathlib import Path


def request(url, token, payload=None):
    data = None if payload is None else json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, headers={'Authorization': 'Bearer '+token, 'Content-Type':'application/json'})
    with urllib.request.urlopen(req, timeout=8) as response:
        return json.load(response)


def select_display(config):
    """Only enrolled endpoints; priority indicates room preference, not distance."""
    for target in sorted(config['displays'], key=lambda d:d['priority']):
        try:
            token = os.environ[target['token_env']]
            health = request(target['url']+'/health', token)
            if health['ready'] and not health['busy'] and target['deck'] in health['decks']:
                return target, token
        except (KeyError, OSError, ValueError):
            continue
    raise RuntimeError('No configured display is ready with the selected deck')


class Robot:
    def __init__(self, config, enabled=False):
        self.config, self.client = config, None
        if enabled:
            if not config.get('commissioned'):
                raise RuntimeError('Robot must be commissioned in config before enabling gestures')
            from unitree_sdk2py.core.channel import ChannelFactoryInitialize
            from unitree_sdk2py.g1.arm.g1_arm_action_client import G1ArmActionClient
            ChannelFactoryInitialize(0, config['interface'])
            self.client = G1ArmActionClient()
            self.client.SetTimeout(5.0)
            self.client.Init()
    def action(self, name):
        action_id = self.config.get('actions', {}).get(name)
        if self.client is None or action_id is None:
            print('Gesture:', name, '(disabled or unmapped)')
            return
        code = self.client.ExecuteAction(int(action_id))
        if code != 0:
            raise RuntimeError(f'G1 action failed: {code}')


def run_session(config, robot, name='visitor'):
    target, token = select_display(config)
    url = target['url']
    print('Selected display:', target['name'])
    robot.action('greet')
    result = request(url+'/start', token, {'deck':target['deck'], 'greeting':config['greeting'].format(name=name)})
    session_id = result['session_id']
    last_slide = 0
    try:
        while True:
            status = request(url+'/status', token)
            if status.get('session_id') != session_id:
                raise RuntimeError('Display session changed unexpectedly')
            if status['state'] == 'error':
                raise RuntimeError(status.get('error', 'Display failed'))
            if status['state'] in ('done','stopped'):
                break
            slide = status.get('slide',0)
            if slide and slide != last_slide:
                robot.action('present')
                last_slide = slide
            time.sleep(0.4)
    finally:
        try:
            request(url+'/stop', token, {'session_id':session_id})
        except Exception as exc:
            print('Remote stop failed:', exc)
    robot.action('finish')


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--config',default='config.json')
    parser.add_argument('--start',action='store_true',help='Manual start without camera')
    parser.add_argument('--enable-robot',action='store_true')
    args=parser.parse_args()
    config=json.loads(Path(args.config).read_text())
    robot=Robot(config['robot'],args.enable_robot)
    if args.start:
        run_session(config,robot)
        return
    import cv2
    detector=cv2.CascadeClassifier(cv2.data.haarcascades+'haarcascade_frontalface_default.xml')
    if detector.empty(): raise RuntimeError('Face cascade failed to load')
    camera=cv2.VideoCapture(config['camera'])
    if not camera.isOpened(): raise RuntimeError('Camera unavailable')
    # Optional locally enrolled LBPH identities. Unknown people still get a generic greeting.
    recognizer=None
    labels={}
    model=Path(config.get('identity_dir','identities'))
    if (model/'model.yml').exists():
        recognizer=cv2.face.LBPHFaceRecognizer_create()
        recognizer.read(str(model/'model.yml'))
        labels=json.loads((model/'labels.json').read_text())
    stable_since=None
    absence_since=None
    armed=True
    try:
        while True:
            ok,frame=camera.read()
            if not ok: raise RuntimeError('Camera frame acquisition failed')
            gray=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
            faces=detector.detectMultiScale(gray,1.1,5,minSize=(80,80))
            now=time.monotonic()
            if len(faces):
                absence_since=None
                if stable_since is None: stable_since=now
                if armed and now-stable_since>=config['face_dwell_seconds']:
                    name='visitor'
                    if recognizer:
                        x,y,w,h=max(faces,key=lambda f:f[2]*f[3])
                        label,distance=recognizer.predict(cv2.resize(gray[y:y+h,x:x+w],(160,160)))
                        if distance<config.get('identity_threshold',55): name=labels.get(str(label),'visitor')
                    armed=False
                    try: run_session(config,robot,name)
                    except Exception as exc: print('Session failed:',exc)
                    stable_since=None
            else:
                stable_since=None
                if absence_since is None: absence_since=now
                if now-absence_since>=config['rearm_absence_seconds']: armed=True
            for x,y,w,h in faces: cv2.rectangle(frame,(x,y),(x+w,y+h),(0,200,0),2)
            cv2.imshow('G1 visitor detection - Q quits',frame)
            if cv2.waitKey(30)&255==ord('q'): break
    finally:
        camera.release()
        cv2.destroyAllWindows()

if __name__=='__main__':
    try: main()
    except KeyboardInterrupt: print('Presenter stopped')
