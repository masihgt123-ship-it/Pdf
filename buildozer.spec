[app]

title = PDF to Word
package.name = pdfword
package.domain = org.pdftools

source.dir = .
source.include_exts = py,kv,png,jpg,jpeg,atlas
source.exclude_dirs = .git,.github,.buildozer,bin,venv,.venv

version = 0.1.0

requirements = python3==3.11.5,hostpython3==3.11.5,kivy,pyjnius,pypdf==5.1.0,python-docx==1.1.2,lxml==4.9.3,typing_extensions

orientation = portrait
fullscreen = 0

android.api = 35
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a

android.accept_sdk_license = True
android.enable_androidx = True

[buildozer]

log_level = 2
warn_on_root = 1
