/**
 * Contains Windows related exceptions classes.
 *
 * Copyright: (c) 2009 John Chapman
 *
 * License: See $(LINK2 ..\..\..\licence.txt, licence.txt) for use and distribution terms.
 */

module mswin.exception;

import win32.windows;
import std.string, std.utf;
import core.exception;

/// Converts an LRESULT/HRESULT error code to a corresponding Exception object.
/// Throwns an exception with a specific failure LRESULT/HRESULT value.
void checkResult(int resultCode)
{
    if(SUCCEEDED(resultCode)) return;

    switch (resultCode) {
        case E_NOTIMPL:
            throw new NotImplementedException;
        case E_NOINTERFACE:
            throw new InvalidCastException;
        case E_POINTER:
            throw new NullReferenceException;
        case E_ACCESSDENIED:
            throw new UnauthorizedAccessException;
        case E_OUTOFMEMORY:
            throw new OutOfMemoryError;
        case E_INVALIDARG:
            throw new ArgumentException;
        default:
    }
    throw new ComException(resultCode);
}

/**
* The exception thrown when one of the arguments provided to a method is not valid.
*/
class ArgumentException : Exception {

    private static const E_ARGUMENT = "Value does not fall within the expected range.";

    private string paramName_;

    this() {
        super(E_ARGUMENT);
    }

    this(string message) {
        super(message);
    }

    this(string message, string paramName) {
        super(message);
        paramName_ = paramName;
    }

    final string paramName() {
        return paramName_;
    }

}

/**
* The exception thrown when a null reference is passed to a method that does not accept it as a valid argument.
*/
class ArgumentNullException : ArgumentException {

    private static const E_ARGUMENTNULL = "Value cannot be null.";

    this() {
        super(E_ARGUMENTNULL);
    }

    this(string paramName) {
        super(E_ARGUMENTNULL, paramName);
    }

    this(string paramName, string message) {
        super(message, paramName);
    }

}

/**
* The exception that is thrown when the value of an argument passed to a method is outside the allowable range of values.
*/
class ArgumentOutOfRangeException : ArgumentException {

    private static const E_ARGUMENTOUTOFRANGE = "Index was out of range.";

    this() {
        super(E_ARGUMENTOUTOFRANGE);
    }

    this(string paramName) {
        super(E_ARGUMENTOUTOFRANGE, paramName);
    }

    this(string paramName, string message) {
        super(message, paramName);
    }

}

/**
* The exception thrown when the format of an argument does not meet the parameter specifications of the invoked method.
*/
class FormatException : Exception {

    private static const E_FORMAT = "The value was in an invalid format.";

    this() {
        super(E_FORMAT);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown for invalid casting.
*/
class InvalidCastException : Exception {

    private static const E_INVALIDCAST = "Specified cast is not valid.";

    this() {
        super(E_INVALIDCAST);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when a method call is invalid.
*/
class InvalidOperationException : Exception {

    private static const E_INVALIDOPERATION = "Operation is not valid.";

    this() {
        super(E_INVALIDOPERATION);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when a requested method or operation is not implemented.
*/
class NotImplementedException : Exception {

    private static const E_NOTIMPLEMENTED = "The operation is not implemented.";

    this() {
        super(E_NOTIMPLEMENTED);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when an invoked method is not supported.
*/
class NotSupportedException : Exception {

    private static const E_NOTSUPPORTED = "The specified method is not supported.";

    this() {
        super(E_NOTSUPPORTED);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when there is an attempt to dereference a null reference.
*/
class NullReferenceException : Exception {

    private static const E_NULLREFERENCE = "Object reference not set to an instance of an object.";

    this() {
        super(E_NULLREFERENCE);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when the operating system denies access.
*/
class UnauthorizedAccessException : Exception {

    private static const E_UNAUTHORIZEDACCESS = "Access is denied.";

    this() {
        super(E_UNAUTHORIZEDACCESS);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when a security error is detected.
*/
class SecurityException : Exception {

    private static const E_SECURITY = "Security error.";

    this() {
        super(E_SECURITY);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown for errors in an arithmetic, casting or conversion operation.
*/
class ArithmeticException : Exception {

    private static const E_ARITHMETIC = "Overflow or underflow in arithmetic operation.";

    this() {
        super(E_ARITHMETIC);
    }

    this(string message) {
        super(message);
    }

}

/**
* The exception thrown when an arithmetic, casting or conversion operation results in an overflow.
*/
class OverflowException : ArithmeticException {

    private const E_OVERFLOW = "Arithmetic operation resulted in an overflow.";

    this() {
        super(E_OVERFLOW);
    }

    this(string message) {
        super(message);
    }

}

static string getErrorMessage(int errorCode) {
    wchar[256] buffer;
    uint result = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, null, errorCode, 0, buffer.ptr, buffer.length + 1, null);
    if (result != 0) {
        string s = toUTF8(buffer[0 .. result]);

        while (result > 0) {
            char c = s[result - 1];
            if (c > ' ' && c != '.')
                break;
            result--;
        }
        return std.string.format("%s. (0x%08X)", s[0 .. result], cast(uint)errorCode);
    }
    return std.string.format("Unspecified error (0x%08X)", errorCode);
}

/**
* Generic Win32 Exception.
*
* The default contructor obtains the last error code and message.
*/
class Win32Exception : Exception {

    private uint errorCode_;

    this(uint errorCode = GetLastError()) {
        this(errorCode, getErrorMessage(errorCode));
    }

    this(uint errorCode, string message) {
        super(message);
        errorCode_ = errorCode;
    }

    @property uint errorCode() {
        return errorCode_;
    }

}

/**
* The exception thrown when an unrecognized HRESULT is returned from a COM operation.
*/
class ComException : Win32Exception {

    /**
    * Initializes a new instance with a specified error code.
    * Params: errorCode = The error code (HRESULT) value associated with this exception.
    */
    this(int errorCode) {
        super(errorCode);
    }
}

// Runtime DLL support.

class DllNotFoundException : Exception {

    this(string message = "Dll was not found.") {
        super(message);
    }

}

class EntryPointNotFoundException : Exception {

    this(string message = "Entry point was not found.") {
        super(message);
    }

}



