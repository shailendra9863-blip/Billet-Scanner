[app]

(str) Title of your application
title = Billet Scanner

(str) Package name
package.name = billetscanner

(str) Package domain (needed for android packaging)
package.domain = org.ccm

(str) Source code where the main.py lives
source.dir = .

(list) Source files to include (let empty to include all the files)
source.include_exts = py,png,jpg,kv,atlas,json

(list) Application requirements
Aapke main.py ke saare external libraries aur unke core dependencies yahan list hain
requirements = python3,kivy==2.3.1,openpyxl,et_xmlfile,requests,urllib3,chardet,idna,certifi

(str) Supported orientations (landscape, portrait or all)
orientation = portrait

(bool) Indicate if the application should be fullscreen or not
fullscreen = 0

=============================================================================
Android specific configuration
=============================================================================

(list) Permissions required by the app
Internet aur Mobile storage mein Excel file save karne ke liye permissions
android.permissions = INTERNET, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, MOUNT_UNMOUNT_FILESYSTEMS

(int) Target Android API (Android 13+)
android.api = 33

(int) Minimum API your APK will support (Android 5.0+)
android.minapi = 21

(str) Android NDK version to use
android.ndk = 25b

(bool) If True, then skip signup for Google Play Quality Protection
android.skip_source_filter = 1

(str) Format used to package the app for debug mode (apk)
android.debug_artifact = apk

=============================================================================
Buildozer settings
=============================================================================

[buildozer]

(int) Log level (0 = error only, 1 = info, 2 = debug)
log_level = 2

(int) Display warning if buildozer is run as root
warn_on_root = 1