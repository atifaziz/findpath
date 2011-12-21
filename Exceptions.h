// FINDPATH - Locates a file using the Windows search path
// Copyright (C) 2004, Atif Aziz http://www.raboof.com
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

#pragma once

// --------------------------------------------------------------------------
//  Exception
// --------------------------------------------------------------------------

class Exception {};

// --------------------------------------------------------------------------
//  ApplicationException
// --------------------------------------------------------------------------

class ApplicationException : public Exception
{
public:

    ApplicationException(LPCTSTR message) : m_message(message) 
    { _ASSERT(message); }

    LPCTSTR GetMessage() const { return m_message; }

private:

    LPCTSTR m_message;

    ApplicationException& operator=(const ApplicationException&);
};

// --------------------------------------------------------------------------
//  SystemException
// --------------------------------------------------------------------------

class SystemException : public Exception
{
public:

    SystemException(DWORD code) : m_code(code) { _ASSERT(NO_ERROR != code); }

    static void ThrowLast() { throw SystemException(GetLastError()); }

    DWORD GetCode() const { return m_code; }

private:

    DWORD m_code;
};

// --------------------------------------------------------------------------
//  ComException
// --------------------------------------------------------------------------

class ComException : public Exception
{
public:

    ComException(HRESULT hr) : m_hr(hr) {}

    static void Check(HRESULT hr) 
    { 
        if (FAILED(hr))
            throw ComException(hr); 
    }

    HRESULT GetCode() const { return m_hr; }

private:

    DWORD m_hr;
};

inline static void CoCheckError(HRESULT hr) 
{ 
    ComException::Check(hr);
}
