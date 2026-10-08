import argparse,json,os
p=argparse.ArgumentParser();p.add_argument('--devices',action='store_true');a=p.parse_args()
if a.devices:
    from screeninfo import get_monitors
    for i,m in enumerate(get_monitors()): print('Monitor',i,m)
    import win32com.client
    voice=win32com.client.Dispatch('SAPI.SpVoice');outputs=voice.GetAudioOutputs()
    for i in range(outputs.Count): print('Audio',outputs.Item(i).GetDescription())
else:
    from presenter import select_display
    c=json.load(open('config.json'));target,_=select_display(c);print('Ready display:',target['name'])
