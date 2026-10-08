"""Enroll consenting visitors locally; run one person at a time."""
import argparse,json,time
from pathlib import Path
import cv2
import numpy as np
p=argparse.ArgumentParser();p.add_argument('name');p.add_argument('--camera',type=int,default=0);p.add_argument('--directory',default='identities');a=p.parse_args()
root=Path(a.directory);root.mkdir(exist_ok=True)
labels=json.loads((root/'labels.json').read_text()) if (root/'labels.json').exists() else {}
label=next((int(k) for k,v in labels.items() if v==a.name),max([int(k) for k in labels]+[-1])+1)
samples=root/str(label);samples.mkdir(exist_ok=True)
cap=cv2.VideoCapture(a.camera);cascade=cv2.CascadeClassifier(cv2.data.haarcascades+'haarcascade_frontalface_default.xml')
if not cap.isOpened(): raise RuntimeError('Camera unavailable')
try:
    for n in range(40):
        ok,frame=cap.read()
        if not ok: raise RuntimeError('Camera read failed')
        gray=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY);faces=cascade.detectMultiScale(gray,1.1,5,minSize=(80,80))
        if len(faces)!=1: raise RuntimeError('Exactly one visible face is required; rerun enrollment')
        x,y,w,h=faces[0];cv2.imwrite(str(samples/f'{time.time_ns()}.png'),cv2.resize(gray[y:y+h,x:x+w],(160,160)))
        cv2.imshow('Enrollment - vary head angle slightly',frame);cv2.waitKey(250)
finally: cap.release();cv2.destroyAllWindows()
labels[str(label)]=a.name
images=[];ids=[]
for folder in root.iterdir():
    if folder.is_dir() and folder.name.isdigit():
        for file in folder.glob('*.png'):
            images.append(cv2.imread(str(file),0));ids.append(int(folder.name))
recognizer=cv2.face.LBPHFaceRecognizer_create();recognizer.train(images,np.array(ids,dtype=np.int32));recognizer.write(str(root/'model.yml'))
(root/'labels.json').write_text(json.dumps(labels,indent=2))
print('Enrollment complete. Recognition is approximate and is not authentication.')
