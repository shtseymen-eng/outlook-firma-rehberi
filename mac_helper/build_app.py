"""Build a macOS launcher for the current architecture and package the local web UI. Run npm ci first."""
import os
from pathlib import Path
import plistlib
import platform
import shutil
import subprocess
import sys

root = Path(__file__).resolve().parent.parent
out = Path(sys.argv[1]).resolve()
app = out/'Mac Mail Rehberi.app'
if app.exists(): shutil.rmtree(app)
macos=app/'Contents/MacOS'; resources=app/'Contents/Resources'
macos.mkdir(parents=True); resources.mkdir(parents=True)
subprocess.run(['npm','run','build'],cwd=root/'web',env={**os.environ,'LOCAL_HELPER':'1'},check=True)
shutil.copytree(root/'web/dist',resources/'web',ignore=shutil.ignore_patterns('Mac-Mail-Rehberi.zip'))
for name in ['server.py','mail_reader.js']: shutil.copy2(root/'mac_helper'/name,resources/name)
cache=root/'build/swift-cache'; cache.mkdir(parents=True,exist_ok=True)
binaries=[]
for arch in [platform.machine()]:
    binary=root/'build'/('launcher-'+arch)
    subprocess.run(['swiftc','-O','-module-cache-path',str(cache),'-target',arch+'-apple-macosx12.0',str(root/'mac_helper/Launcher.swift'),'-o',str(binary)],check=True)
    binaries.append(str(binary))
subprocess.run(['lipo','-create',*binaries,'-output',str(macos/'MacMailRehberi')],check=True)
plist={'CFBundleExecutable':'MacMailRehberi','CFBundleIdentifier':'com.shtseymen.macmailrehberi','CFBundleName':'Mac Mail Rehberi','CFBundlePackageType':'APPL','CFBundleShortVersionString':'1.0.0','CFBundleVersion':'1','LSMinimumSystemVersion':'12.0','NSHighResolutionCapable':True,'NSAppleEventsUsageDescription':'Mail hesaplarınızdaki gönderen ve alıcı adreslerini yerel firma rehberine aktarmak için yalnızca okuma erişimi gerekir.'}
(app/'Contents/Info.plist').write_bytes(plistlib.dumps(plist))
for key in ['com.apple.FinderInfo','com.apple.ResourceFork']:
    subprocess.run(['xattr','-r','-d',key,str(app)],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
subprocess.run(['codesign','--force','--deep','--sign','-',str(app)],check=True)
archive=out/'Mac-Mail-Rehberi.zip'
subprocess.run(['ditto','-c','-k','--norsrc','--noextattr','--keepParent',str(app),str(archive)],check=True)
print('Built:',app,'Archive:',archive)
