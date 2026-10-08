# G1 humanoid meet and greet presentation package

This package implements a stationary visitor-triggered presenter. The controller detects a face, optionally recognizes a locally enrolled visitor, gives a configured G1 greeting gesture, chooses a ready enrolled display, starts native PowerPoint slideshow mode, and narrates speaker notes through the TV speaker. The display worker changes slides only after speech finishes. Default robot mode prints gesture events without moving the robot.

The supplied ref12.docx was 30,323 zero-filled bytes, not a readable DOCX. Its content could not be incorporated. This implementation follows the written request. Robo G1 is assumed to mean Unitree G1; confirm the manufacturer and model before installing the robot SDK. Real robot, camera, Windows PowerPoint, HDMI and audio acceptance tests remain to be performed on your equipment.

## Architecture

Camera -> presenter.py on controller -> authenticated HTTP -> display_agent.py on Windows PC -> native PowerPoint -> HDMI or paired wireless TV.

presenter.py -> Unitree SDK2 DDS -> G1 greeting and configured presentation gestures.

The controller can run on a Windows laptop in robot-disabled mode or on a Linux companion PC with the Unitree SDK. The Windows PC can be a small PC mounted behind the TV. Do not assume the robot's onboard computer runs Windows or PowerPoint. Use a USB webcam first; the G1 camera requires a working vendor camera stream, supplied to config.camera as an OpenCV-compatible source. This package does not implement an undocumented G1 camera driver.

## What works and what must be commissioned

| Feature | Implementation | Setup required |
|---|---|---|
| Visitor detection | OpenCV frontal face cascade and 2 second dwell | Camera aimed at visitors |
| Name recognition | Optional LBPH local enrollment | enroll.py, consent and threshold validation |
| Automatic display selection | Health probe of configured endpoints, priority and deck availability | Enroll each room endpoint; priorities express room preference |
| Native PPT slideshow | PowerPoint COM worker | Windows and installed desktop PowerPoint |
| Text narration | Offline Windows SAPI, speaker notes then slide text | Installed language voice; selected TV audio endpoint |
| Remote display | HTTP worker running on display PC | Same trusted LAN or secured VPN |
| Screen sharing | Windows display projection or meeting app sharing | First pairing and explicit audio sharing |
| G1 gesture | Official SDK2 ExecuteAction | Supported G1 variant, interface, posture and action commissioning |

Automatic selection means selecting an available configured endpoint, not measuring physical distance or pairing arbitrary nearby TVs. For location awareness, set room-specific priorities or later connect room tags/beacons to endpoint selection. Unsupported displays fail visibly; they are not silently claimed as connected.

## Step 1 Install on the Windows display PC

Use Python 3.12 and an interactive signed-in Windows desktop. Install Microsoft desktop PowerPoint. PowerPoint COM cannot be exercised on Linux and unattended Windows services are unsuitable for this worker. Existing Windows/Office licensing applies; the Python libraries themselves are open source.

Extract the ZIP. Open PowerShell in the extracted G1_Presenter directory:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install python-pptx
.\.venv\Scripts\python.exe make_demo.py
.\.venv\Scripts\python.exe diagnostics.py --devices
```

For your own deck, copy presentation.pptx into this folder and change display_config.json decks.demo to its absolute path. Write what the robot should say in each slide's speaker notes. Without notes the worker reads text boxes. It does not describe images, charts or nested SmartArt automatically. Expand those descriptions into speaker notes. Disable automatic transitions, embedded autoplay media and animation builds for the first run; the worker narrates entire slides and does not step through animation builds.

## Step 2 Connect HDMI and TV audio

Connect the display PC HDMI to the TV. Select the corresponding TV input. Press Win+P and select Extend. Set monitor_index in display_config.json to the TV index printed by diagnostics.py. Leave null only if you have confirmed PowerPoint's configured default slideshow monitor is correct.

Open Windows Settings -> System -> Sound -> Output and select the TV/HDMI device. Verify a test sound. diagnostics.py --devices lists the SAPI outputs. Set G1_AUDIO_DEVICE to a unique substring of the exact HDMI output name if SAPI lists it. Some Windows installations expose only the default endpoint; then leave G1_AUDIO_DEVICE unset and set the Windows default output to the TV before starting the worker. Restart the worker after any audio device change. No separate paid TTS service is needed. English and other languages depend on installed SAPI voices; this worker uses the current default voice.

Generate and record a random token:

```powershell
.\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
$env:G1_DISPLAY_TOKEN = 'paste-generated-token-here'
# Optional when a uniquely matching SAPI device exists:
$env:G1_AUDIO_DEVICE = 'TV output name'
.\.venv\Scripts\python.exe display_agent.py
```

Leave this terminal open. In a second PowerShell terminal in the same folder, set the same token and run:

```powershell
$env:G1_DISPLAY_TOKEN = 'same-token'
.\.venv\Scripts\python.exe diagnostics.py
.\.venv\Scripts\python.exe presenter.py --start
```

You should see the greeting, three slides and synchronized audio. Ctrl+C in the controller requests session stop; the worker purges speech and closes its slideshow. Use the authenticated /stop endpoint from another terminal if the controller loses connectivity. The robot's hardware emergency stop remains separate from presentation stop.

## Step 3 Enable visitor detection

```powershell
.\.venv\Scripts\python.exe presenter.py
```

A face present continuously for two seconds triggers one presentation. Visitors remaining in view do not repeatedly restart it. The camera must observe no faces for ten seconds to rearm. Q closes detection while idle; Ctrl+C cancels during a presentation. Adjust camera, face_dwell_seconds and rearm_absence_seconds in config.json.

Optional personalized greeting:

```powershell
.\.venv\Scripts\python.exe enroll.py Nani
.\.venv\Scripts\python.exe presenter.py
```

Enrollment captures 40 local samples. Enroll in representative lighting and test known and unknown visitors. LBPH similarity is approximate, can misidentify faces and provides no liveness check. A detected or recognized face never authorizes robot motion beyond the previously commissioned gesture set. Delete identities/ to remove all enrollment data. Generic face detection stores no face images.

## Step 4 Remote TV or screen sharing

For a TV in another room, attach a Windows PC to it and install this display worker there. On that PC change host in display_config.json to its LAN interface IP, e.g. 192.168.1.50. Set its token_env to G1_REMOTE_TOKEN and set that environment variable. Allow inbound TCP 8765 in Windows Firewall only from the controller's IP. Copy the PPTX onto that PC and configure its local deck path and TV monitor index.

On the controller set G1_REMOTE_TOKEN to the same token and update the remote URL in config.json. Assign the remote endpoint priority 1 if it should be preferred; otherwise the ready local endpoint wins. The controller transfers commands, not presentation files. All selected display PCs must already have the configured deck. HTTP bearer tokens are for a trusted lab LAN; use a VPN or TLS reverse proxy outside it and do not expose port 8765 publicly.

For Miracast, enable the receiver's wireless display mode, press Win+K on the Windows PC, select the receiver and complete pairing. Use Win+P Extend and select the newly enumerated monitor and audio output as above. Once paired and connected, the worker treats it like another Windows monitor. Reconnection/pairing is performed by Windows, not this package.

For Teams/Zoom or another screenshare app, join the meeting on the display PC, explicitly share the slideshow screen/window and enable system audio sharing. Start presenter.py afterward. Meeting admission, credentials and screenshare clicks are handled in the meeting app; this package does not silently control them. Remote worker mode is preferable for unattended LAN demonstration because it avoids screenshare audio and admission problems.

## Step 5 Commission Unitree G1 gestures

Official sources:
- https://github.com/unitreerobotics/unitree_sdk2_python
- https://github.com/unitreerobotics/unitree_sdk2_python/blob/master/unitree_sdk2py/g1/arm/g1_arm_action_client.py
- https://github.com/unitreerobotics/unitree_sdk2_python/blob/master/unitree_sdk2py/g1/audio/g1_audio_client.py

On the supported Linux companion controller install the official SDK according to its current README and CycloneDDS requirements. Record the tested SDK commit and robot firmware. For example after prerequisites:

```bash
git clone https://github.com/unitreerobotics/unitree_sdk2_python.git
python -m pip install -e ./unitree_sdk2_python
```

Use the vendor-supported wired robot network and determine the correct interface with ip link. Set robot.interface accordingly; eth0 is only an example. Configure the display URL to the Windows TV PC IP, set its token in the controller shell, install requirements, and confirm manual presentation works with robot disabled.

The verified official arm API exposes G1ArmActionClient.Init, ExecuteAction(action_id), and GetActionList. The current SDK source maps face wave to 25. Treat that mapping as firmware-dependent and confirm the robot's supported action list with the vendor example before enabling it. The package maps only greet=25 by default; present and finish are null until you map suitable verified actions. There is no locomotion, autonomous approach, turning, low-level joint trajectory or navigation code here.

Commission with a trained operator, the vendor-prescribed stable posture/support arrangement, clear arm space, and the hardware emergency stop available. The application does not infer safe posture from a face. Once verified, set robot.commissioned=true, and run:

```bash
export G1_REMOTE_TOKEN='same-display-token'
python presenter.py --enable-robot
```

This dispatches greet once per session and any configured present action once per observed slide change. Network polling means gestures are approximately synchronized, not real-time motion choreography. SDK success means command acceptance, not verified physical completion. Keep present/finish disabled until validated; do not use repeated gestures whose duration overlaps slide changes. TV narration is the implemented audio route. The official G1 AudioClient has TtsMaker/PlayStream, but this package does not substitute robot-speaker narration because speech-completion feedback differs.

## Acceptance procedure

1. With robot disabled, run tests and the manual demo. Confirm native full-screen slides on the intended TV and no narration from laptop speakers.
2. Add speaker notes to a five-slide real deck. Check text-to-slide alignment and a slide containing only images with explanatory notes.
3. Trigger with an unknown visitor, then an enrolled visitor. Test sustained presence, absence and re-entry. Check false recognition under your lighting.
4. Stop the highest-priority worker before starting: verify fallback to the next configured ready endpoint. If none is ready, the controller reports an error. During an active presentation it aborts on lost connectivity; it does not restart on another screen mid-deck.
5. Ctrl+C during speech: confirm narration stops and the worker-owned slideshow exits. Test a missing deck, wrong token, busy worker and disconnected HDMI.
6. Commission a single greeting on your G1. Then validate a complete visitor session with the TV. Record firmware, SDK commit, display and camera configuration.

## Validation provided with this package

Run python -m unittest -v. Automated tests cover display priority/fallback, no ready display, robot commissioning gate, remote error cleanup, authentication, busy session rejection, matching stop identity, and cancellable speech. They use mocks; they do not prove Windows COM, TV connectivity, recognition accuracy or robot behavior. Python syntax and ZIP integrity are checked separately.

## Troubleshooting

- Worker ready=false: inspect display worker console/status; check SAPI output and Windows COM dependencies.
- PowerPoint failure: start it manually once, finish first-run dialogs and open the deck. Run the worker as the same signed-in desktop user.
- Wrong TV: run diagnostics, update monitor_index and restart after display changes.
- Silent TV: verify Windows sound test, volume/mute and HDMI device; restart the worker after changing output.
- Camera unavailable: set another camera index or a working compatible stream; close apps occupying the camera.
- Robot import/DDS error: confirm SDK installation, supported G1 variant, network interface and vendor prerequisites. Continue the display demo in robot-disabled mode while commissioning.
- Corrupt PPTX or missing deck: fix the actual file/path; display agent accepts only configured local PPTX files.

## Export narration audio files

On Windows, with python-pptx installed, run:

```powershell
.\.venv\Scripts\python.exe export_audio.py presentation.pptx --out narration_audio
```

This generates one offline SAPI WAV file per slide using notes, then slide text as fallback. The live display worker synthesizes speech directly instead of replaying these files; exported files are useful for review or manual playback on the TV.
