[app]

title = PDF Tools

package.name = pdfword
package.domain = org.pdftools

source.dir = .

source.include_exts = py,kv,png,jpg,jpeg,atlas,traineddata
source.exclude_dirs = .git,.github,.buildozer,bin,venv,.venv,tests

version = 0.2.1
android.numeric_version = 3

requirements = python3==3.11.5,hostpython3==3.11.5,kivy,pyjnius,pypdf==5.1.0,python-docx==1.1.2,lxml==4.9.3,typing_extensions

orientation = portrait
fullscreen = 0

android.api = 35
android.minapi = 24
android.ndk = 25b

android.archs = arm64-v8a

android.accept_sdk_license = True
android.enable_androidx = True

android.add_src = java

android.gradle_dependencies = com.rmtheis:tess-two:9.1.0

android.release_artifact = apk


[buildozer]

log_level = 2
warn_on_root = 1
