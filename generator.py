# -*- coding: utf-8 -*-


"""
Generate control panel task links: command, name, and keywords

This script is modified from this: https://github.com/blackdaemon/enso-launcher-continued/blob/master/enso/contrib/open/platform/win32/control_panel_win7.py
Need pywin32

Output language will depends on current system's language, not only English

Output:
{
  "schema_version": "0.0.1",
  "language": "Chinese (Simplified)_China",
  "windows_version": "19043.1055",
  "items": [
    {
      "name": "\u8bad\u7ec3\u8ba1\u7b97\u673a\u4ee5\u8bc6\u522b\u4f60\u7684\u58f0\u97f3",
      "cmd": "%%windir%\\system32\\rundll32.exe %windir%\\system32\\speech\\speechux\\SpeechUX.dll,RunWizard UserTraining",
      "keywords": [
        [
          "\u7cbe\u786e\u6027",
          "\u7cbe\u786e\u5730",
          "\u7cbe\u786e",
          "\u7cbe\u5ea6",
          "\u66f4\u6b63",
          "\u9519\u8bef",
          "accuracy",
          "accurately",
          "precise",
          "precision",
          "correct",
          "mistakes"
        ],
        [
          "\u8ba1\u7b97\u673a",
          "\u673a\u5668",
          "\u6211\u7684",
          "\u4e2a\u4eba",
          "\u7cfb\u7edf",
          "\u7535\u8111",
          "\u6211\u7684\u7535\u8111",
          "compuer",
          "machine",
          "my",
          "computer",
          "pc",
          "personal",
          "system"
        ],
        [
          "\u81ea\u5b9a\u4e49",
          "\u81ea\u5b9a",
          "\u4e2a\u6027\u5316",
          "\u4e2a\u6027",
          "customisation",
          "customises",
          "customization",
          "customizes",
          "customizing",
          "personalisation",
          "personalise",
          "personalization",
          "personalize"
        ],
...
"""
import datetime
import json
import locale
import os
import re
import winreg
from collections import OrderedDict
from xml.etree import cElementTree as ElementTree

import win32api
import win32con


def _expand_path_variables(file_path):
    re_env = re.compile(r'%\w+%')

    def expander(mo):
        return os.environ.get(mo.group()[1:-1], 'UNKNOWN')

    return re_env.sub(expander, file_path)


def read_mui_string_from_dll(id):
    assert id.startswith("@") and ",-" in id, \
        "id has invalid format. Expected format is '@dllfilename,-id'"

    m = re.match("@([^,]+),-([0-9]+)(;.*)?", id)
    if m:
        dll_filename = _expand_path_variables(m.group(1))
        string_id = int(m.group(2))  # python 3 style
    else:
        raise Exception(
            "Error parsing dll-filename and string-resource-id from '%s'" % id)

    h = win32api.LoadLibraryEx(
        dll_filename,
        None,
        win32con.LOAD_LIBRARY_AS_DATAFILE | win32con.DONT_RESOLVE_DLL_REFERENCES)
    if h:
        s = win32api.LoadString(h, string_id)
        return s
    return None


def read_task_links_xml(xml):
    APPS_NS = "http://schemas.microsoft.com/windows/cpltasks/v1"
    TASKS_NS = "http://schemas.microsoft.com/windows/tasks/v1"
    TASKS_NS2 = "http://schemas.microsoft.com/windows/tasks/v2"

    results = []
    tree = ElementTree.fromstring(xml)

    for app in tree.findall("{%s}application" % APPS_NS):
        for item in app.findall('{%s}task' % TASKS_NS):
            keyword_tags = item.findall('{%s}keywords' % TASKS_NS)
            keywords = []
            for m in [k.text for k in keyword_tags]:
                keywords.append([kw.strip() for kw in
                                 read_mui_string_from_dll(m).split(';') if
                                 kw.strip()])

            name_shell32id = item.findtext("{%s}name" % TASKS_NS)
            name = read_mui_string_from_dll(name_shell32id)

            cmd = item.findtext("{%s}command" % TASKS_NS)

            # This specific task link has an extra Percent Sign, not sure why, only this one has, bug of MsWindows?
            # "name": "Train the computer to recognise your voice",
            # "cmd": "%%windir%\\system32\\rundll32.exe %windir%\\system32\\speech\\speechux\\SpeechUX.dll,RunWizard UserTraining",
            if name == "Train the computer to recognise your voice" and cmd.startswith("%%"):
                cmd = cmd[1:]

            cp = item.find("{%s}controlpanel" % TASKS_NS2)

            if cmd is not None:
                pass
            elif cp is not None:
                cname = cp.get('name')
                cpage = cp.get('page')
                cmd = "%%SystemRoot%%\\System32\\control.exe -name %s" % cname
                if cpage:
                    cmd += " /page %s" % cpage
            else:
                raise ValueError('cmd and cp both None')
            results.append((name, cmd, keywords))

    return results


def read_default_windows7_tasklist_xml():
    handle = win32api.LoadLibraryEx(
        "shell32.dll",
        None,
        win32con.LOAD_LIBRARY_AS_DATAFILE | win32con.DONT_RESOLVE_DLL_REFERENCES)
    # HARD-CODED!, Is the resource-id #21 for all versions of Windows 7?
    xml = win32api.LoadResource(handle, "XML", 21)
    return xml


def generate():
    final = OrderedDict()
    final['schema_version'] = "0.0.1"
    final['language'] = locale.getlocale()[0]

    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"

    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as key:
        current_build_number = winreg.QueryValueEx(key, 'CurrentBuildNumber')[0]
        ubr = winreg.QueryValueEx(key, 'UBR')[0]
    final['windows_version'] = f'{current_build_number}.{ubr}'

    xml = read_default_windows7_tasklist_xml()
    r = read_task_links_xml(xml)
    list_item = []
    for idx, i in enumerate(r):
        list_item.append(
            {
                'name': i[0],
                'cmd': i[1],
                'keywords': i[2],
            }
        )

    final['items'] = sorted(list_item, key=lambda i: i['cmd'])

    with open(f"{datetime.datetime.now().strftime('%Y%b%dT%H%M%S')}"
              f"-{final['schema_version']}"
              f"-{final['windows_version']}"
              f"-{final['language']}.json",
              'wt', encoding='utf8') as f_out:
        json.dump(final, f_out, indent=2)


if __name__ == '__main__':
    generate()
