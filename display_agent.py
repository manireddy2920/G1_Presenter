"""Authenticated LAN PowerPoint worker. Run on the Windows PC connected to the TV."""
import argparse, json, os, queue, secrets, threading, time
from pathlib import Path
from flask import Flask, request, jsonify

app=Flask(__name__)
state={'state':'idle','slide':0,'ready':False,'busy':False}
lock=threading.Lock()
jobs=queue.Queue()
stop=threading.Event()
TOKEN=''
DECKS={}
MONITOR=None

def update(**values):
    with lock: state.update(values)

@app.before_request
def auth():
    supplied=request.headers.get('Authorization','')
    if not TOKEN or not secrets.compare_digest(supplied,'Bearer '+TOKEN):
        return jsonify(error='Unauthorized'),401

@app.get('/health')
def health():
    with lock: return jsonify(ready=state['ready'],busy=state['busy'],decks=list(DECKS))

@app.get('/status')
def status():
    with lock: return jsonify(dict(state))

@app.post('/start')
def start():
    body=request.get_json(silent=True) or {}
    deck=body.get('deck')
    if deck not in DECKS: return jsonify(error='Unknown configured deck'),400
    greeting=body.get('greeting','Welcome. Let us begin.')
    if not isinstance(greeting,str) or len(greeting)>400: return jsonify(error='Invalid greeting'),400
    with lock:
        if not state['ready']: return jsonify(error='Worker not ready'),503
        if state['busy']: return jsonify(error='Presentation in progress'),409
        sid=secrets.token_hex(12)
        state.update(busy=True,state='queued',session_id=sid,slide=0,error=None)
        stop.clear()
        jobs.put((deck,greeting,sid))
    return jsonify(session_id=sid)

@app.post('/stop')
def halt():
    body=request.get_json(silent=True) or {}
    with lock:
        if body.get('session_id') != state.get('session_id'):
            return jsonify(error='Session mismatch'),409
        stop.set()
    return jsonify(ok=True)


def speak(voice,text):
    if stop.is_set(): return
    voice.Speak(text,1) # SAPI async; poll for true completion rather than word-count timers.
    while not voice.WaitUntilDone(100):
        if stop.is_set():
            voice.Speak('',3) # Async + purge pending utterances
            return


def worker():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        voice=win32com.client.Dispatch('SAPI.SpVoice')
        configured=os.environ.get('G1_AUDIO_DEVICE','')
        outputs=voice.GetAudioOutputs()
        if configured:
            matches=[outputs.Item(i) for i in range(outputs.Count) if configured.lower() in outputs.Item(i).GetDescription().lower()]
            if len(matches)!=1: raise RuntimeError('G1_AUDIO_DEVICE must match exactly one SAPI output')
            voice.AudioOutput=matches[0]
        print('Speech output:',voice.AudioOutput.GetDescription())
        update(ready=True)
        while True:
            deck,greeting,sid=jobs.get()
            ppt=pres=show=None
            try:
                # Own a dedicated instance and close only this worker's presentation.
                ppt=win32com.client.DispatchEx('PowerPoint.Application')
                pres=ppt.Presentations.Open(str(DECKS[deck]),ReadOnly=True,Untitled=False,WithWindow=False)
                settings=pres.SlideShowSettings
                settings.RangeType=1
                settings.ShowType=1
                settings.AdvanceMode=1 # Manual navigation; narration controls slide changes.
                show=settings.Run()
                if MONITOR is not None:
                    from screeninfo import get_monitors
                    monitors=get_monitors()
                    if MONITOR>=len(monitors): raise RuntimeError('Configured monitor no longer exists')
                    m=monitors[MONITOR]
                    import win32gui, win32con
                    win32gui.SetWindowPos(show.HWND,win32con.HWND_TOP,m.x,m.y,m.width,m.height,win32con.SWP_SHOWWINDOW)
                update(state='greeting')
                speak(voice,greeting)
                for i in range(1,pres.Slides.Count+1):
                    if stop.is_set(): break
                    show.View.GotoSlide(i)
                    update(state='presenting',slide=i)
                    slide=pres.Slides.Item(i)
                    notes=[]
                    for shape in slide.NotesPage.Shapes:
                        if shape.Type==14 and shape.PlaceholderFormat.Type==2 and shape.HasTextFrame:
                            notes.append(shape.TextFrame.TextRange.Text)
                    text=' '.join(notes).strip()
                    if not text:
                        text=' '.join(shape.TextFrame.TextRange.Text for shape in slide.Shapes if shape.HasTextFrame).strip()
                    speak(voice,text or f'Slide {i}.')
                    stop.wait(0.7)
                update(state='stopped' if stop.is_set() else 'done')
            except Exception as exc:
                update(state='error',error=str(exc))
            finally:
                for obj,method in ((show.View if show else None,'Exit'),(pres,'Close'),(ppt,'Quit')):
                    if obj:
                        try: getattr(obj,method)()
                        except Exception: pass
                update(busy=False)
                jobs.task_done()
    except Exception as exc:
        update(ready=False,state='error',error=str(exc))
    finally: pythoncom.CoUninitialize()


def main():
    global TOKEN,DECKS,MONITOR
    parser=argparse.ArgumentParser()
    parser.add_argument('--config',default='display_config.json')
    args=parser.parse_args()
    cfg=json.loads(Path(args.config).read_text())
    TOKEN=os.environ.get(cfg['token_env'],'')
    if len(TOKEN)<24: raise RuntimeError('Set a random token of at least 24 characters')
    DECKS={key:Path(path).resolve() for key,path in cfg['decks'].items()}
    for p in DECKS.values():
        if not p.is_file() or p.suffix.lower()!='.pptx': raise RuntimeError(f'Missing PPTX: {p}')
    MONITOR=cfg.get('monitor_index')
    threading.Thread(target=worker,daemon=True).start()
    from waitress import serve
    serve(app,host=cfg.get('host','127.0.0.1'),port=cfg.get('port',8765),threads=4)

if __name__=='__main__': main()
