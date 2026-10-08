import unittest,os
from unittest.mock import patch
import presenter
import display_agent as agent
class Tests(unittest.TestCase):
 def test_priority_and_unavailable(self):
  cfg={'displays':[{'priority':2,'url':'http://b','token_env':'TEST_TOKEN','deck':'demo','name':'b'},{'priority':1,'url':'http://a','token_env':'TEST_TOKEN','deck':'demo','name':'a'}]}
  with patch.dict(os.environ,{'TEST_TOKEN':'secret'}),patch('presenter.request',side_effect=[OSError(),{'ready':True,'busy':False,'decks':['demo']}]) as call:
   self.assertEqual(presenter.select_display(cfg)[0]['name'],'b');self.assertIn('http://a',call.call_args_list[0].args[0])
 def test_no_display(self):
  with self.assertRaises(RuntimeError): presenter.select_display({'displays':[]})
 def test_robot_gate(self):
  with self.assertRaises(RuntimeError): presenter.Robot({'commissioned':False},True)
 def test_session_failure_stops(self):
  target={'url':'http://a','name':'a','deck':'demo'}
  robot=presenter.Robot({})
  with patch('presenter.select_display',return_value=(target,'token')),patch('presenter.request',side_effect=[{'session_id':'x'},{'session_id':'x','state':'error','error':'failed'},{'ok':True}]) as call:
   with self.assertRaises(RuntimeError): presenter.run_session({'greeting':'Hello {name}'},robot)
   self.assertEqual(call.call_args.args[0],'http://a/stop')
 def test_agent_auth_busy_stop(self):
  agent.TOKEN='x'*32;agent.DECKS={'demo':'demo.pptx'};agent.state.update(ready=True,busy=False)
  client=agent.app.test_client();headers={'Authorization':'Bearer '+agent.TOKEN}
  self.assertEqual(client.get('/health').status_code,401)
  self.assertEqual(client.post('/start',headers=headers,json={'deck':'missing'}).status_code,400)
  r=client.post('/start',headers=headers,json={'deck':'demo'});self.assertEqual(r.status_code,200)
  self.assertEqual(client.post('/start',headers=headers,json={'deck':'demo'}).status_code,409)
  self.assertEqual(client.post('/stop',headers=headers,json={'session_id':'wrong'}).status_code,409)
  self.assertEqual(client.post('/stop',headers=headers,json={'session_id':r.json['session_id']}).status_code,200)
 def test_speech_cancel(self):
  from unittest.mock import Mock
  voice=Mock();agent.stop.clear()
  def wait(_): agent.stop.set();return False
  voice.WaitUntilDone.side_effect=wait;agent.speak(voice,'text')
  self.assertEqual(voice.Speak.call_args.args,('',3))
if __name__=='__main__':unittest.main()
