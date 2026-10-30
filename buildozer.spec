[app]
title = Sistema de Deudores
package.name = sistemadeudores
package.domain = ar.com.sistemadeudores
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 1.0.0
requirements = python3,kivy,openpyxl
orientation = landscape
fullscreen = 0
android.permissions =
android.api = 35
android.minapi = 24
android.ndk = 28b
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True
p4a.bootstrap = sdl2
p4a.branch = develop
[buildozer]
log_level = 2
warn_on_root = 1
