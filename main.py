#encoding=utf8

from __future__ import division
import json
import os
import collections
import string
import ctypes
import sys
from wox import Wox,WoxAPI

class space(Wox):

    def bytes2human(self,n):
        symbols = ('KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB')
        prefix = {}
        for i, s in enumerate(symbols):
            prefix[s] = 1 << (i+1)*10
        for s in reversed(symbols):
            if n >= prefix[s]:
                value = float(n) / prefix[s]
                return '%.1f%s' % (value, s)
        return "%sB" % n

    def get_drives(self):
        kernel32 = ctypes.windll.kernel32
        GetVolumeInformationW = kernel32.GetVolumeInformationW
        GetDiskFreeSpaceExW = kernel32.GetDiskFreeSpaceExA
        bitmask = kernel32.GetLogicalDrives()
        GetDriveType = kernel32.GetDriveTypeA

        def volume_name(path):
            volumeNameBuffer = ctypes.create_unicode_buffer(1024)
            fileSystemNameBuffer = ctypes.create_unicode_buffer(1024)
            serial_number = None
            max_component_length = None
            file_system_flags = None

            GetVolumeInformationW(
                ctypes.c_wchar_p(path),
                volumeNameBuffer,
                ctypes.sizeof(volumeNameBuffer),
                serial_number,
                max_component_length,
                file_system_flags,
                fileSystemNameBuffer,
                ctypes.sizeof(fileSystemNameBuffer)
            )

            return "No label" if volumeNameBuffer.value == "" else volumeNameBuffer.value

        def disk_usage(path):
            _, total, free = ctypes.c_ulonglong(), ctypes.c_ulonglong(), \
                               ctypes.c_ulonglong()
            ret = GetDiskFreeSpaceExW(path, ctypes.byref(_), ctypes.byref(total), ctypes.byref(free))
            return [total.value, total.value - free.value, free.value]

        drives = []
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                if GetDriveType(letter+':/') == 3  or GetDriveType(letter+':/') == 2 :
                    du =  disk_usage(letter+':/')
                    du.insert(0, letter)
                    du.insert(1, volume_name(letter+':/'))
                    du.insert(2, GetDriveType(letter+':/'))
                    drives.append( du )
            bitmask >>= 1

        return drives

    def query(self,key):
        results = []
        drives = self.get_drives()
        for drive in drives:
            path = drive[0]+':/'
            label = drive[1]
            icon = "Images/hdd.ico" if drive[2] == 3 else "Images/rdd.ico"
            total = drive[3]
            used = drive[4]
            free = drive[5]
            percent = int(round(free/total*100))
            title = label + ': ' + str(percent)+'% ('+self.bytes2human(free)+') free'
            SubTitle = self.bytes2human(total - free) + ' used of ' + self.bytes2human(total) + ' total'

            results.append({"Title": title, "SubTitle": SubTitle, "IcoPath":icon,"JsonRPCAction":{"method": "openDrive","parameters":[path],"dontHideAfterAction":False}})

        return results

    def openDrive(self,path):
        os.startfile(path)

if __name__ == "__main__":
    space()