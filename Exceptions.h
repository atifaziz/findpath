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
