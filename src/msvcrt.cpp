/* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1/GPL 2.0/LGPL 2.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is for Open Layer for Unicode (opencow).
 *
 * The Initial Developer of the Original Code is Brodie Thiesfield.
 * Portions created by the Initial Developer are Copyright (C) 2004
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *
 * Alternatively, the contents of this file may be used under the terms of
 * either the GNU General Public License Version 2 or later (the "GPL"), or
 * the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
 * in which case the provisions of the GPL or the LGPL are applicable instead
 * of those above. If you wish to allow use of your version of this file only
 * under the terms of either the GPL or the LGPL, and not to allow others to
 * use your version of this file under the terms of the MPL, indicate your
 * decision by deleting the provisions above and replace them with the notice
 * and other provisions required by the GPL or the LGPL. If you do not delete
 * the provisions above, a recipient may use your version of this file under
 * the terms of any one of the MPL, the GPL or the LGPL.
 *
 * ***** END LICENSE BLOCK ***** */

// define these symbols so that we don't get dllimport linkage 
// from the system headers
#define _CRTIMP

#include <windows.h>
#include <wchar.h>
#include <direct.h>
#include <crtdbg.h>
#include <stdio.h>
#include <io.h>
#include <stdarg.h>
#include <sys/types.h>
#include <sys/stat.h>

#include "MbcsBuffer.h"

// ----------------------------------------------------------------------------
// API

int __cdecl 
_wrename(
    const wchar_t * oldname, 
    const wchar_t * newname
    )
{
    CMbcsBuffer mbcsOldname;
    if (!mbcsOldname.FromUnicode(oldname)) {
        errno = ENOENT;
        return -1;
    }

    CMbcsBuffer mbcsNewname;
    if (!mbcsNewname.FromUnicode(newname)) {
        errno = ENOENT;
        return -1;
    }

    return ::rename(mbcsOldname, mbcsNewname);
}

int __cdecl 
_wremove(
    const wchar_t * path
    )
{
    CMbcsBuffer mbcsPath;
    if (!mbcsPath.FromUnicode(path)) {
        errno = ENOENT;
        return -1;
    }

    return ::remove(mbcsPath);
}

int __cdecl 
_wopen(
    const wchar_t * filename, 
    int oflag, 
    ... /* [int pmode] */
    )
{
    va_list marker;
    int pmode;

    va_start(marker, oflag);
    pmode = va_arg(marker, int);
    va_end(marker);

    CMbcsBuffer mbcsFilename;
    if (!mbcsFilename.FromUnicode(filename)) {
        errno = ENOENT;
        return -1;
    }

    return ::_open(mbcsFilename, oflag, pmode);
}

int __cdecl 
_waccess(
    const wchar_t * path, 
    int mode
    )
{
    CMbcsBuffer mbcsPath;
    if (!mbcsPath.FromUnicode(path)) {
        errno = ENOENT;
        return -1;
    }

    return ::_access(mbcsPath, mode);
}

int __cdecl 
_wchdir(
    const wchar_t * dirname
    )
{
    CMbcsBuffer mbcsDirname;
    if (!mbcsDirname.FromUnicode(dirname)) {
        errno = ENOENT;
        return -1;
    }

    return ::_chdir(mbcsDirname);
}

wchar_t * __cdecl 
_wgetcwd(
    wchar_t * buffer, 
    int maxlen
    )
{
    CMbcsBuffer mbcsBuffer;
    if (!::_getcwd(mbcsBuffer, mbcsBuffer.BufferSize()))
        return NULL;

    int nBufferLen = ::lstrlenA(mbcsBuffer) + 1;
    int nRequiredLen = ::MultiByteToWideChar(CP_ACP, 0, mbcsBuffer, nBufferLen, 0, 0);
    if (buffer && maxlen < nRequiredLen) {
        errno = ERANGE;
        return NULL;
    }

    if (!buffer) {
        if (maxlen < nRequiredLen)
            maxlen = nRequiredLen;
        buffer = (wchar_t *) malloc(sizeof(wchar_t) * maxlen);
        if (!buffer) {
            errno = ENOMEM;
            return NULL;
        }
    }

    ::MultiByteToWideChar(CP_ACP, 0, mbcsBuffer, nBufferLen, buffer, maxlen);
    return buffer;
}

wchar_t * __cdecl 
_wgetdcwd( 
   int drive,
   wchar_t *buffer,
   int maxlen
   )
{
    char * localBuffer = NULL;
    char stackBuffer[MAX_PATH];
    char * cwd;
    wchar_t * outputBuffer;
    int byteLen, charLen, outputBufferLen;

    // use a stack buffer if possible, otherwise we let
    // the CRT allocate it for us
    if (buffer && maxlen <= MAX_PATH)
        localBuffer = stackBuffer;

    // get the current drive path, this will allocate the buffer if
    // localBuffer is still set to NULL. If there is an error we can 
    // just return because we haven't allocated anything and the errno
    // will be set appropriately.
    cwd = ::_getdcwd(drive, localBuffer, localBuffer ? MAX_PATH : 0);
    if (!cwd)
        return NULL;
    byteLen = (int) ::lstrlenA(cwd) + 1; // byte length including null terminator

    // determine the size of the buffer required and return the appropriate
    // error if they have not supplied a buffer big enough
    charLen = ::MultiByteToWideChar(CP_ACP, 0, cwd, byteLen, 0, 0);
    if (charLen == 0 || (buffer && maxlen < charLen)) {
        if (cwd != localBuffer)
            ::free(cwd);
        errno = ERANGE;
        return NULL;
    }

    // allocate an output buffer if required
    outputBuffer = buffer;
    outputBufferLen = maxlen;
    if (!outputBuffer) {
        if (outputBufferLen < charLen)
            outputBufferLen = charLen;
        outputBuffer = (wchar_t*) ::malloc(outputBufferLen * sizeof(wchar_t));
        if (!outputBuffer) {
            if (cwd != localBuffer)
                ::free(cwd);
            errno = ENOMEM;
            return NULL;
        }
    }

    // convert the MBCS path to wide char, we already know that this
    // will not fail because we definitely have a buffer big enough
    // and it has already scanned the string for bad characters.
    ::MultiByteToWideChar(CP_ACP, 0, cwd, byteLen, outputBuffer, outputBufferLen);
    if (cwd != localBuffer)
        ::free(cwd);
    return outputBuffer;
}

int __cdecl 
_wmkdir(
    const wchar_t * dirname
    )
{
    CMbcsBuffer mbcsDirname;
    if (!mbcsDirname.FromUnicode(dirname)) {
        errno = ENOENT;
        return -1;
    }

    return ::_mkdir(mbcsDirname);
}

int __cdecl 
_wrmdir(
    const wchar_t * dirname
    )
{
    CMbcsBuffer mbcsDirname;
    if (!mbcsDirname.FromUnicode(dirname)) {
        errno = ENOENT;
        return -1;
    }

    return ::_rmdir(mbcsDirname);
}

int __cdecl 
_wstat(
    const wchar_t * path, 
    struct _stat * buffer
    )
{
    CMbcsBuffer mbcsPath;
    if (!mbcsPath.FromUnicode(path)) {
        errno = ENOENT;
        return -1;
    }

    return ::_stat(mbcsPath, buffer);
}

int __cdecl 
_wstati64(
    const wchar_t * path, 
    struct _stati64 * buffer
    )
{
    CMbcsBuffer mbcsPath;
    if (!mbcsPath.FromUnicode(path)) {
        errno = ENOENT;
        return -1;
    }

    return ::_stati64(mbcsPath, buffer);
}

