/**
* Provides additional support for COM (Component Object Model).
*
* Copyright: (c) 2009 John Chapman
*
* License: See $(LINK2 ..\..\..\licence.txt, licence.txt) for use and distribution terms.
*/

module mswin.automation;

public import mswin.com;
import mswin.registry;

import win32.oaidl, win32.wtypes, win32.ocidl;

import core.vararg;
import core.stdc.string : memcpy;

import std.string;
import std.algorithm;
import std.exception;
import std.traits;
import std.conv;

pragma(lib, "oleaut32.lib");


/**
* Encapsulate a BSTR for automatic conversion/cleanup
*/
struct ComStr
{
    BSTR _bstr; // a wchar*
    alias _bstr this;

    this(this)
    {
        // make a copy. we were blitted
        _bstr = SysAllocString(_bstr);
    }

    this(S)(S str) {
        this = str; // forward to opAssign, should be just a pointer anyway
    }

    void opAssign(S)(S str)  
    {
        clear();

        static if(isSomeString!S)
            _bstr = SysAllocString(std.utf.toUTF16z(str));
        else static if(isPointer!S && isSomeChar!(typeof(*S.init)))
            _bstr = SysAllocString(std.utf.toUTF16z(str[0 .. zeroLen(str)]));
        else static if (is(S == ComStr))
            swap(_bstr, str._bstr);
        else
            static assert(false);
    }

    ~this()
    {
        clear();
    }

    string toString()
    {
        return cast(string)(this);
    }

    auto opCast(C)()
    {  
        if(_bstr is null)
            return null;

        uint len = SysStringLen(_bstr);

        return to!C(_bstr[0 .. len]);
    }

    /**
    * Returns the address of the hold BSTR. Use this for [out] functions.
    * The previous content will be cleared.
    */
    @property BSTR* pptr()
    {
        clear();
        return &_bstr;
    }

    void clear()
    {
        if(_bstr !is null) {
            SysFreeString(_bstr);
            _bstr = null;
        }
    }

    /// Attaches to an existing BSTR, taking ownership
    void attach(BSTR str)
    {
        clear();
        _bstr = str;
    }

    /// Detaches from the hold BSTR
    BSTR detach()
    {
        BSTR val = _bstr;
        _bstr = null;
        return val;
    }

}

/**
* Converts a BSTR to a string, optionally freeing the original BSTR.
* Params: bstr = The BSTR to convert.
* Returns: A string equivalent to bstr.
*/
string fromBstr(wchar* s, bool free = true) {
    if (s == null)
        return null;

    uint len = SysStringLen(s);
    if (len == 0)
        return null;

    string ret = std.utf.toUTF8(s[0 .. len]);
    if (free)
        SysFreeString(s);
    return cast(string)ret;
}

size_t zeroLen(C)(const(C)* ptr)
{
    size_t len = 0;

    while(*ptr != '\0')
    {
        ++ptr;
        ++len;
    }

    return len;
}

/**
* Encapsulate a VARIANT, container for many different types.
* Examples:
* ---
* ComVariant var = 10;     // Instance contains VT_I4.
* var = "Hello, World"; // Instance now contains VT_BSTR.
* var = 234.5;          // Instance now contains VT_R8.
* ---
*/
struct ComVariant {

    VARIANT _v;
    alias _v this;

    static ComVariant Missing() 
    { 
        static ComVariant v; 
        v._v.vt = VARENUM.VT_ERROR; v._v.scode = DISP_E_PARAMNOTFOUND; 
        return v; 
    };

    static ComVariant Nothing()
    { 
        static ComVariant v;
        v._v.vt = VARENUM.VT_DISPATCH;
        v._v.pdispVal = null;
        return v;
    }

    static ComVariant Null() 
    { 
        static ComVariant v;
        v._v.vt = VARENUM.VT_NULL;
        return v;
    }

    /**
    * Initializes a new instance using the specified value.
    * A GUID will create the object and query it for IDsipatch.
    */
    this(T)(T value) {
        static if (is(T E == enum)) {
            return opCall(cast(E)value, VariantType!(T));
        }
        else static if(is(T == GUID))
        {
            auto obj = ComPtr!IDispatch(value);
            this = obj;
        }
        else {
            VARTYPE type = VariantType!(T);
            this = value;
            if (type != vt)
                VariantChangeType(&_v, &_v, VARIANT_ALPHABOOL, type).checkResult();
        }
    }

    void opAssign(T)(T value) {
        if (vt != VARENUM.VT_EMPTY)
            clear();

        static if (is(T == VARIANT_BOOL)) boolVal = value;
        else static if (is(T == bool)) boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
        else static if (is(T == ubyte)) bVal = value;
        else static if (is(T == byte)) cVal = value;
        else static if (is(T == ushort)) uiVal = value;
        else static if (is(T == short)) iVal = value;
        else static if (is(T == uint)) ulVal = value;
        else static if (is(T == int)) lVal = value;
        else static if (is(T == ulong)) ullVal = value;
        else static if (is(T == long)) llVal = value;
        else static if (is(T == float)) fltVal = value;
        else static if (is(T == double)) dblVal = value;
        else static if (is(T == DECIMAL)) decVal = value;
        else static if (is(T : string)) bstrVal = ComStr(value).detach();
        else static if (is(T : IDispatch)) pdispVal = value, value.AddRef();
        else static if (is(T : IUnknown)) punkVal = value, value.AddRef();
        else static if (is(T : Object)) byref = cast(void*)value;
        else static if (is(T == VARIANT*)) pvarVal = value;
        else static if (is(T == VARIANT)) VariantCopy(&_v, &value);
        else static if (is(T == ComVariant)) swap(this, value);
        else static if (is(T == SAFEARRAY*)) parray = value;
        else static if (isArray!(T)) parray = SAFEARRAY.from(value);
        else static assert(false, "'" ~ T.stringof ~ "' is not one of the allowed types.");

        //if value is a VARIANT, then VariantCopy will have set the type already.
        static if (!is(T == VARIANT) && !is(T == ComVariant))
            vt = VariantType!(T);

        static if (is(T == SAFEARRAY*)) {
            VARTYPE type;
            SafeArrayGetVartype(value, type);
            vt |= type;
        }
    }

    T opCast(T)() {
        const type = VariantType!(T);

        static if (type != VARENUM.VT_VOID) {
            ComVariant temp = this;
            temp.changeType(type);
            with (temp) {
                static if (type == VARENUM.VT_BOOL) {
                    static if (is(T == bool))
                        return (boolVal == VARIANT_TRUE) ? true : false;
                    else
                        return boolVal;
                }
                else static if (type == VARENUM.VT_UI1) return bVal;
                else static if (type == VARENUM.VT_I1) return cVal;
                else static if (type == VARENUM.VT_UI2) return uiVal;
                else static if (type == VARENUM.VT_I2) return iVal;
                else static if (type == VARENUM.VT_UI4) return ulVal;
                else static if (type == VARENUM.VT_I4) return lVal;
                else static if (type == VARENUM.VT_UI8) return ullVal;
                else static if (type == VARENUM.VT_I8) return llVal;
                else static if (type == VARENUM.VT_R4) return fltVal;
                else static if (type == VARENUM.VT_R8) return dblVal;
                else static if (type == VARENUM.VT_DECIMAL) return decVal;
                else static if (type == VARENUM.VT_BSTR) {
                    static if (is(T : string))
                        return fromBstr(bstrVal);
                    else
                        return bstrVal;
                }
                else static if (type == VARENUM.VT_UNKNOWN) return punkVal;
                else static if (type == VARENUM.VT_DISPATCH) return pdispVal;
                else return T.init;
            }
        }
        else static assert(false, "Cannot cast from '" ~ U.stringof ~ "' to '" ~ T.stringof ~ "'.");
    }

    this(this)
    {
        // We were postblited already. Check if we need a real copy
        if( vt == VT_BSTR || 
           vt == VT_ARRAY || 
           vt == VT_VARIANT || 
           vt == VT_DISPATCH || 
           vt == VT_UNKNOWN || 
           vt & VT_ARRAY || 
           vt & VARENUM.VT_BYREF || 
           vt == VT_VOID )
        {
            // make a new copy of it
            VARIANT temp;
            VariantCopy(&temp, &_v);
            _v = temp;
        }
    }

    ~this()
    {
        clear();
    }

    /**
    * Clears the value of this instance and releases any associated memory.
    * See_Also: $(LINK2 http://msdn2.microsoft.com/en-us/library/ms221165.aspx, VariantClear).
    */
    void clear() {
        if (!(isNull || isEmpty))
            VariantClear(&_v);
    }

    /**
    * Copies this instance into the destination value.
    * Params: dest = The variant to copy into.
    */
    void copyTo(out VARIANT dest) {
        VariantCopy(&dest, &_v).checkResult();
    }

    /**
    * Convers this variant from one type to another.
    * Params: newType = The type to change to.
    */
    void changeType(VARTYPE newType) {
        VariantChangeTypeEx(&_v, &_v, GetThreadLocale(), VARIANT_ALPHABOOL, newType).checkResult();
    }

    /**
    * Converts the value contained in this instance to a string.
    * Returns: A string representation of the value contained in this instance.
    */
    string toString() {
        if (isNull || isEmpty)
            return null;

        if (vt == VARENUM.VT_BSTR)
            return fromBstr(bstrVal, false);

        VARIANT temp;
        VariantChangeType(&temp, &_v, VARIANT_ALPHABOOL | VARIANT_LOCALBOOL, VARENUM.VT_BSTR).checkResult();
        return fromBstr(temp.bstrVal);
    }

    /**
    * Returns the value contained in this instance.
    * The requested type must match the current type in this variant.
    * If you need conversion use to!T
    */
    @property V value(V)() {
        enforce(VariantType!V == vt); 
        static if (is(V == long)) return llVal;
        else static if (is(V == int)) return lVal;
        else static if (is(V == ubyte)) return bVal;
        else static if (is(V == short)) return iVal;
        else static if (is(V == float)) return fltVal;
        else static if (is(V == double)) return dblVal;
        else static if (is(V == bool)) return (boolVal == VARIANT_TRUE) ? true : false;
        else static if (is(V == VARIANT_BOOL)) return boolVal;
        else static if (is(V : string)) return fromBstr(bstrVal, false);
        else static if (is(V == wchar*)) return bstrVal;
        else static if (is(V : IDispatch)) return cast(V)pdispVal;
        else static if (is(V : IUnknown)) return cast(V)punkVal;
        else static if (is(V == SAFEARRAY*)) return parray;
        else static if (isArray!(V)) return parray.toArray!(typeof(*V))();
        else static if (is(V == VARIANT*)) return pvarVal;
        else static if (is(V : Object)) return cast(V)byref;
        else static if (isPointer!(V)) return cast(V)byref;
        else static if (is(V == byte)) return cVal;
        else static if (is(V == ushort)) return uiVal;
        else static if (is(V == uint)) return ulVal;
        else static if (is(V == ulong)) return ullVal;
        else static if (is(V == DECIMAL)) return decVal;
        else static assert(false, "'" ~ V.stringof ~ "' is not one of the allowed types.");
    }

    enum : VARIANT_BOOL {
        True = VARIANT_TRUE,
        False = VARIANT_FALSE
    }

    /**
    * Determines whether this instance is empty.
    */
    @property bool isEmpty() {
        return (vt == VARENUM.VT_EMPTY);
    }

    /**
    * Determines whether this instance is _null.
    */
    @property bool isNull() {
        return (vt == VARENUM.VT_NULL);
    }

    /**
    * Determines whether this instance is Nothing.
    */
    @property bool isNothing() {
        return (vt == VARENUM.VT_DISPATCH && pdispVal is null)
            || (vt == VARENUM.VT_UNKNOWN && punkVal is null);
    }

    int opCmp(ComVariant that) {
        return VarCmp(&_v, &that._v, GetThreadLocale(), 0) - 1;
    }

    bool opEquals(ComVariant that) {
        return opCmp(that) == 0;
    }

    /// As we have alias this and opDispatch, we need to forward the VARIANT members
    auto opDispatch(string name)() if(is(typeof(mixin("_v." ~ name))))
    {
        mixin("return _v." ~ name ~ ";" );
    }

    /** IDispatch automation helpers.
    * Examples:
    * ----
    *   auto excel = ComVariant(clsidFromProgId("Excel.Application"));
    *   excel.put_Visible(true); // put property 'Visible'
    *   auto workbook = excel.get_Workbooks.Add(); // chaining
    * ----
    */
    auto opDispatch(string name, A...)(auto ref A args) if(!is(typeof(mixin("_v." ~ name))))
    {
        enforce(vt == VT_DISPATCH, "can not invoke on this type");

        static if(name.startsWith("get_")) {
            return getProperty(pdispVal, name[4 .. $], args);
        }
        else static if(name.startsWith("put_")) {
            return setProperty(pdispVal, name[4 .. $], args);
        }
        else 
            return invokeMethod(pdispVal, name, args);

        assert(false);
    }

}

/**
* Determines whether this instance is empty.
*/
bool isEmpty(VARIANT v) {
    return (v.vt == VARENUM.VT_EMPTY);
}

/**
* Determines whether this instance is _null.
*/
bool isNull(VARIANT v) {
    return (v.vt == VARENUM.VT_NULL);
}

/**
* Determines whether this instance is Nothing.
*/
bool isNothing(VARIANT v) {
    return (v.vt == VARENUM.VT_DISPATCH && v.pdispVal is null)
        || (v.vt == VARENUM.VT_UNKNOWN && v.punkVal is null);
}

/**
* Determines the equivalent COM type of a built-in type at compile-time.
* Examples:
* ---
* auto a = VariantType!(string);          // VT_BSTR
* auto b = VariantType!(bool);            // VT_BOOL
* auto c = VariantType!(typeof([1,2,3])); // VT_ARRAY | VT_I4
* ---
*/
template VariantType(T) {
    static if (is(T == VARIANT_BOOL))
        const VariantType = VARENUM.VT_BOOL;
    else static if (is(T == bool))
        const VariantType = VARENUM.VT_BOOL;
    else static if (is(T == char))
        const VariantType = VARENUM.VT_UI1;
    else static if (is(T == ubyte))
        const VariantType = VARENUM.VT_UI1;
    else static if (is(T == byte))
        const VariantType = VARENUM.VT_I1;
    else static if (is(T == ushort))
        const VariantType = VARENUM.VT_UI2;
    else static if (is(T == short))
        const VariantType = VARENUM.VT_I2;
    else static if (is(T == uint))
        const VariantType = VARENUM.VT_UI4;
    else static if (is(T == int))
        const VariantType = VARENUM.VT_I4;
    else static if (is(T == ulong))
        const VariantType = VARENUM.VT_UI8;
    else static if (is(T == long))
        const VariantType = VARENUM.VT_I8;
    else static if (is(T == float))
        const VariantType = VARENUM.VT_R4;
    else static if (is(T == double))
        const VariantType = VARENUM.VT_R8;
    else static if (is(T == DECIMAL))
        const VariantType = VARENUM.VT_DECIMAL;
    else static if (is(T E == enum))
        const VariantType = VariantType!(E);
    else static if (is(T : string) || is(T : wstring) || is(T : dstring))
        const VariantType = VARENUM.VT_BSTR;
    else static if (is(T == wchar*))
        const VariantType = VARENUM.VT_BSTR;
    else static if (is(T == SAFEARRAY*))
        const VariantType = VARENUM.VT_ARRAY;
    else static if (is(T == VARIANT))
        const VariantType = VARENUM.VT_VARIANT;
    else static if (is(T : IDispatch))
        const VariantType = VARENUM.VT_DISPATCH;
    else static if (is(T : IUnknown))
        const VariantType = VARENUM.VT_UNKNOWN;
    else static if (isArray!(T))
        const VariantType = VariantType!(typeof(*T)) | VARENUM.VT_ARRAY;
    else static if (isPointer!(T)/* && !is(T == void*)*/)
        const VariantType = VariantType!(typeof(*T)) | VARENUM.VT_BYREF;
    else
        const VariantType = VT_VOID;
}

unittest {
    ComVariant pvarLeft, pvarRight, pvarResult;

    VariantInit(pvarLeft);
    VariantClear(pvarLeft);
    VariantCopy(pvarLeft, pvarRight);
    VarAdd(pvarLeft, pvarRight, pvarResult);
    VarAnd(pvarLeft, pvarRight, pvarResult);
    VarCat(pvarLeft, pvarRight, pvarResult);
    VarDiv(pvarLeft, pvarRight, pvarResult);
    VarMod(pvarLeft, pvarRight, pvarResult);
    VarMul(pvarLeft, pvarRight, pvarResult);
    VarOr(pvarLeft, pvarRight, pvarResult);
    VarSub(pvarLeft, pvarRight, pvarResult);
    VarXor(pvarLeft, pvarRight, pvarResult);
    VarCmp(pvarLeft, pvarRight, GetThreadLocale(), 0);
}

/// Encapsulate a SAFEARRAY* for automatic memory management.
struct ComSafeArray 
{
    SAFEARRAY* _sa;
    alias _sa this;

    this(this)
    {
        if(_sa is null) return;
        // we were bit copyied. make a real copy.
        SAFEARRAY* sa;
        SafeArrayCopy(_sa, &sa).checkResult();
        _sa = sa;
    }

    ~this()
    {
        clear();
    }

    /**
    * Initializes a new instance using the specified _array.
    * Params: array = The elements with which to initialize the instance.
    * Returns: A pointer to the new instance.
    */
    static SAFEARRAY* opCall(T)(T[] array) {
        auto bound = SAFEARRAYBOUND(array.length);
        auto sa = SafeArrayCreate(VariantType!(T), 1, &bound);

        static if (is(T : string)) alias wchar* Type;
        else                       alias T Type;

        Type* data;
        SafeArrayAccessData(sa, retval(data));
        for (auto i = 0; i < array.length; i++) {
            static if (is(T : string)) data[i] = ComStr(array[i]).detach();
            else                       data[i] = array[i];
        }
        SafeArrayUnaccessData(sa);

        return sa;
    }

    /**
    * Copies the elements of the SAFEARRAY to a new array of the specified type.
    * Returns: An array of the specified type containing copies of the elements of the SAFEARRAY.
    */
    T[] toArray(T)() {
        int upperBound, lowerBound;
        SafeArrayGetUBound(this, 1, upperBound);
        SafeArrayGetLBound(this, 1, lowerBound);
        int count = upperBound - lowerBound + 1;

        if (count == 0) return null;

        T[] result = new T[count];

        static if (is(T : string)) alias wchar* Type;
        else                       alias T Type;

        Type* data;
        SafeArrayAccessData(this, retval(data));
        for (auto i = lowerBound; i < upperBound + 1; i++) {
            static if (is(T : string)) result[i] = fromBstr(data[i]);
            else                       result[i] = data[i];
        }
        SafeArrayUnaccessData(this);

        return result;
    }

    /**
    * Destroys the SAFEARRAY and all of its data.
    * Remarks: If objects are stored in the array, Release is called on each object.
    */
    void clear() {
        if(_sa !is null) {
            SafeArrayDestroy(_sa);
            _sa = null;
        }
    }

    /**
    * Increments the _lock count of an array.
    */
    void lock() {
        SafeArrayLock(_sa);
    }

    /**
    * Decrements the lock count of an array.
    */
    void unlock() {
        SafeArrayUnlock(_sa);
    }

    /**
    * Gets or sets the number of elements in the array.
    * Params: value = The number of elements.
    */
    void length(int value) {
        auto bound = SAFEARRAYBOUND(value);
        SafeArrayRedim(_sa, &bound);
    }
    /// ditto
    int length() {
        int upperBound, lowerBound;

        SafeArrayGetUBound(_sa, 1, &upperBound);
        SafeArrayGetLBound(_sa, 1, &lowerBound);

        return upperBound - lowerBound + 1;
    }

}

///////////////////////////////////////////////////////////
// IDispatch helpers. Method invocation, property get/set
///////////////////////////////////////////////////////////

/// Specifies the type of member to that is to be invoked.
enum DispatchFlags : ushort {
    InvokeMethod   = DISPATCH_METHOD,         /// Specifies that a method is to be invoked.
    GetProperty    = DISPATCH_PROPERTYGET,    /// Specifies that the value of a property should be returned.
    PutProperty    = DISPATCH_PROPERTYPUT,    /// Specifies that the value of a property should be set.
    PutRefProperty = DISPATCH_PROPERTYPUTREF  /// Specifies that the value of a property should be set by reference.
}

/// The exception thrown when there is an attempt to dynamically access a member that does not exist.
class MissingMemberException : Exception {

    private const string E_MISSINGMEMBER = "Member not found.";

    this() {
        super(E_MISSINGMEMBER);
    }

    this(string message) {
        super(message);
    }

    this(string className, string memberName) {
        super("Member '" ~ className ~ "." ~ memberName ~ "' not found.");
    }

}

/**
* Invokes the specified member on the specified object.
* Params:
*   dispId = The identifier of the method or property member to invoke.
*   flags = The type of member to invoke.
*   target = The object on which to invoke the specified member.
*   args = A list containing the arguments to pass to the member to invoke.
* Returns: The return value of the invoked member.
* Throws: ComException if the call failed.
*/
ComVariant invokeMemberById(int dispId, DispatchFlags flags, IDispatch target, ComVariant[] args...) {
    args.reverse;

    DISPPARAMS params;
    if (args.length > 0) {
        params.rgvarg = cast(VARIANT*)args.ptr;
        params.cArgs = cast(uint)args.length;

        if (flags & DispatchFlags.PutProperty) {
            int dispIdNamed = DISPID_PROPERTYPUT;
            params.rgdispidNamedArgs = &dispIdNamed;
            params.cNamedArgs = 1;
        }
    }

    ComVariant result;
    EXCEPINFO excep;
    int hr = target.Invoke(dispId, &emptyGUID, GetThreadLocale(), cast(ushort)flags, &params, cast(VARIANT*)&result, &excep, null);

    for (auto i = 0; i < params.cArgs; i++) {
        params.rgvarg[i].clear();
    }

    string errorMessage;
    if (hr == DISP_E_EXCEPTION && excep.scode != 0) {
        errorMessage = fromBstr(excep.bstrDescription);
        hr = excep.scode;
    }

    switch (hr) {
        case S_OK, S_FALSE, E_ABORT:
            return result;
        default:
            if (auto supportErrorInfo = ComPtr!ISupportErrorInfo(target)) {
                if (SUCCEEDED(supportErrorInfo.InterfaceSupportsErrorInfo(&uuidof!(IDispatch)))) {
                    IErrorInfo errorInfo;
                    GetErrorInfo(0, &errorInfo);
                    if (errorInfo !is null) {
                        scope(exit) errorInfo.Release();

                        wchar* bstrDesc;
                        if (SUCCEEDED(errorInfo.GetDescription(&bstrDesc)))
                            errorMessage = fromBstr(bstrDesc);
                    }
                }
            }
            else if (errorMessage == null) {
                wchar[256] buffer;
                uint r = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, null, hr, 0, buffer.ptr, buffer.length + 1, null);
                if (r != 0)
                    errorMessage = .toUTF8(buffer[0 .. r]);
                else
                    errorMessage = std.string.format("Operation 0x%08X did not succeed (0x%08X)", dispId, hr);
            }

            throw new ComException(hr);
    }
}

/**
* Invokes the specified member on the specified object.
* Params:
*   name = The _name of the method or property member to invoke.
*   flags = The type of member to invoke.
*   target = The object on which to invoke the specified member.
*   args = A list containing the arguments to pass to the member to invoke.
* Returns: The return value of the invoked member.
* Throws: MissingMemberException if the member is not found.
*/
ComVariant invokeMember(string name, DispatchFlags flags, IDispatch target, ComVariant[] args...) {
    int dispId = DISPID_UNKNOWN;
    ComStr[1] bstrNames = ComStr(name);

    if (SUCCEEDED(target.GetIDsOfNames(&emptyGUID, cast(BSTR*)bstrNames.ptr, 1, GetThreadLocale(), &dispId)) && dispId != DISPID_UNKNOWN) {
        return invokeMemberById(dispId, flags, target, args);
    }

    string typeName;
    ITypeInfo typeInfo;
    if (SUCCEEDED(target.GetTypeInfo(0, 0, &typeInfo))) {
        scope(exit) tryRelease(typeInfo);

        wchar* bstrTypeName;
        typeInfo.GetDocumentation(-1, &bstrTypeName, null, null, null);
        typeName = fromBstr(bstrTypeName);
    }

    throw new MissingMemberException(typeName, name);
}

private ComVariant[] argsToVariantList(TypeInfo[] types, core.vararg.va_list argptr) {
    ComVariant[] list;

    foreach (type; types) {
        if (type == typeid(bool)) list ~= ComVariant(va_arg!(bool)(argptr));
        else if (type == typeid(ubyte)) list ~= ComVariant(va_arg!(ubyte)(argptr));
        else if (type == typeid(byte)) list ~= ComVariant(va_arg!(byte)(argptr));
        else if (type == typeid(ushort)) list ~= ComVariant(va_arg!(ushort)(argptr));
        else if (type == typeid(short)) list ~= ComVariant(va_arg!(short)(argptr));
        else if (type == typeid(uint)) list ~= ComVariant(va_arg!(uint)(argptr));
        else if (type == typeid(int)) list ~= ComVariant(va_arg!(int)(argptr));
        else if (type == typeid(ulong)) list ~= ComVariant(va_arg!(ulong)(argptr));
        else if (type == typeid(long)) list ~= ComVariant(va_arg!(long)(argptr));
        else if (type == typeid(float)) list ~= ComVariant(va_arg!(float)(argptr));
        else if (type == typeid(double)) list ~= ComVariant(va_arg!(double)(argptr));
        else if (type == typeid(string)) list ~= ComVariant(va_arg!(string)(argptr));
        else if (type == typeid(IDispatch)) list ~= ComVariant(va_arg!(IDispatch)(argptr));
        else if (type == typeid(IUnknown)) list ~= ComVariant(va_arg!(IUnknown)(argptr));
        else if (type == typeid(VARIANT)) list ~= ComVariant(va_arg!(VARIANT)(argptr));
        else if (type == typeid(ComVariant)) list ~= va_arg!(ComVariant)(argptr);
        //else if (type == typeid(VARIANT*)) list ~= VARIANT(va_arg!(VARIANT*)(argptr));
        else if (type == typeid(VARIANT*)) list ~= ComVariant(*va_arg!(VARIANT*)(argptr));
    }

    return list;
}

private void fixArgs(ref TypeInfo[] args, ref core.vararg.va_list argptr) {
    if (args[0] == typeid(TypeInfo[]) && args[1] == typeid(core.vararg.va_list)) {
        args = va_arg!(TypeInfo[])(argptr);
        argptr = *cast(core.vararg.va_list*)(argptr);
    }
}

/**
* Invokes the specified method on the specified object.
* Params:
*   target = The object on which to invoke the specified method.
*   name = The _name of the method to invoke.
*   _argptr = A list containing the arguments to pass to the method to invoke.
* Returns: The return value of the invoked method.
* Throws: MissingMemberException if the method is not found.
* Examples:
* ---
* import mswin.com.automation;
*
* void main() {
*   auto ieApp = ComPtr!IDispatch("InternetExplorer.Application");
*   invokeMethod(ieApp, "Navigate", "http://www.amazon.co.uk");
* }
* ---
*/
R invokeMethod(R = ComVariant)(IDispatch target, string name, ...) {
    auto args = _arguments;
    auto argptr = _argptr;
    if (args.length == 2) fixArgs(args, argptr);

    ComVariant ret = invokeMember(name, DispatchFlags.InvokeMethod, target, argsToVariantList(args, argptr));
    static if (is(R == ComVariant)) {
        return ret;
    }
    else {
        return cast(R)ret;
    }
}

/**
* Gets the value of the specified property on the specified object.
* Params:
*   target = The object on which to invoke the specified property.
*   name = The _name of the property to invoke.
*   _argptr = A list containing the arguments to pass to the property.
* Returns: The return value of the invoked property.
* Throws: MissingMemberException if the property is not found.
* Examples:
* ---
* import mswin.com.automation, std.stdio;
*
* void main() {
*   // Create an instance of the Microsoft Word automation object.
*   auto wordApp = ComPtr!IDispatch("Word.Application");
*
*   // Invoke the Documents property
*   //   wordApp.Document
*   ComPtr!IDispatch documents = getProperty!(IDispatch)(target, "Documents");
*
*   // Invoke the Count property on the Documents object
*   //   documents.Count
*   ComVariant count = getProperty(documents, "Count");
*
*   // Display the value of the Count property.
*   writefln("There are %s documents", count);
* }
* ---
*/
R getProperty(R = ComVariant)(IDispatch target, string name, ...) {
    auto args = _arguments;
    auto argptr = _argptr;
    if (args.length == 2) fixArgs(args, argptr);

    ComVariant ret = invokeMember(name, DispatchFlags.GetProperty, target, argsToVariantList(args, argptr));
    static if (is(R == ComVariant))
        return ret;
    else
        return cast(R)ret;
}

/**
* Sets the value of a specified property on the specified object.
* Params:
*   target = The object on which to invoke the specified property.
*   name = The _name of the property to invoke.
*   _argptr = A list containing the arguments to pass to the property.
* Throws: MissingMemberException if the property is not found.
* Examples:
* ---
* import mswin.com.automation;
*
* void main() {
*   // Create an Excel automation object.
*   auto excelApp = ComPtr!IDispatch("Excel.Application");
*
*   // Set the Visible property to true
*   //   excelApp.Visible = true
*   setProperty(excelApp, "Visible", true);
*
*   // Get the Workbooks property
*   //   workbooks = excelApp.Workbook
*   ComPtr!IDispatch workbooks = getProperty!(IDispatch)(excelApp, "Workbooks");
*
*   // Invoke the Add method on the Workbooks property
*   //   newWorkbook = workbooks.Add()
*   ComPtr!IDispatch newWorkbook = invokeMethod!(IDispatch)(workbooks, "Add");
*
*   // Get the Worksheets property and the Worksheet at index 1
*   //   worksheet = excelApp.Worksheets[1]
*   ComPtr!IDispatch worksheet = getProperty!(IDispatch)(excelApp, "Worksheets", 1);
*
*   // Get the Cells property and set the Cell object at column 5, row 3 to a string
*   //   worksheet.Cells[5, 3] = "data"
*   setProperty(worksheet, "Cells", 5, 3, "data");
* }
* ---
*/
void setProperty(IDispatch target, string name, ...) {
    auto args = _arguments;
    auto argptr = _argptr;
    if (args.length == 2) fixArgs(args, argptr);

    if (args.length > 1) {
        ComVariant v = invokeMember(name, DispatchFlags.GetProperty, target);
        if (auto indexer = v.pdispVal) {
            scope(exit) indexer.Release();

            v = invokeMemberById(0, DispatchFlags.GetProperty, indexer, argsToVariantList(args[0 .. 1], argptr));
            if (auto value = v.pdispVal) {
                scope(exit) value.Release();

                invokeMemberById(0, DispatchFlags.PutProperty, value, argsToVariantList(args[1 .. $], argptr + args[0].tsize));
                return;
            }
        }
    }
    else {
        invokeMember(name, DispatchFlags.PutProperty, target, argsToVariantList(args, argptr));
    }
}

/// ditto
void setRefProperty(IDispatch target, string name, ...) {
    auto args = _arguments;
    auto argptr = _argptr;
    if (args.length == 2) fixArgs(args, argptr);

    invokeMember(name, DispatchFlags.PutRefProperty, target, argsToVariantList(args, argptr));
}

/**
*/
class EventCookie(T) {

    private IConnectionPoint cp_;
    private uint cookie_;

    /**
    */
    this(IUnknown source) {
        auto cpc = ComPtr!IConnectionPointContainer(source);
        assert (!cpc.isNull);

        if (cpc.FindConnectionPoint(uuidof!(T), cp_) != S_OK)
            throw new ArgumentException("Source object does not expose '" ~ T.stringof ~ "' event interface.");
    }

    ~this() {
        disconnect();
    }

    /**
    */
    void connect(IUnknown sink) {
        if (cp_.Advise(sink, cookie_) != S_OK) {
            cookie_ = 0;
            tryRelease(cp_);
            throw new InvalidOperationException("Could not Advise() the event interface '" ~ T.stringof ~ "'.");
        }

        if (cp_ is null || cookie_ == 0) {
            if (cp_ !is null)
                tryRelease(cp_);
            throw new ArgumentException("Connection point for event interface '" ~ T.stringof ~ "' cannot be created.");
        }
    }

    /**
    */
    void disconnect() {
        if (cp_ !is null && cookie_ != 0) {
            try {
                cp_.Unadvise(cookie_);
            }
            finally {
                tryRelease(cp_);
                cp_ = null;
                cookie_ = 0;
            }
        }
    }

}

private struct MethodProxy 
{

    int delegate() method;
    VARTYPE returnType;
    VARTYPE[] paramTypes;

    static MethodProxy opCall(R, T...)(R delegate(T) method) {
        MethodProxy self;
        self = method;
        return self;
    }

    void opAssign()(MethodProxy mp) {
        method = mp.method;
        returnType = mp.returnType;
        paramTypes = mp.paramTypes;
    }

    void opAssign(R, T...)(R delegate(T) dg) {
        alias ParameterTypeTuple!(dg) params;

        method = cast(int delegate())dg;
        returnType = VariantType!(R);
        paramTypes.length = params.length;
        foreach (i, paramType; params) {
            paramTypes[i] = VariantType!(paramType);
        }
    }

    int invoke(VARIANT*[] args, VARIANT* result) {

        size_t variantSize(VARTYPE vt) {
            switch (vt) {
                case VARENUM.VT_UI8, VARENUM.VT_I8, VARENUM.VT_CY:
                    return long.sizeof / int.sizeof;
                case VARENUM.VT_R8, VARENUM.VT_DATE:
                    return double.sizeof / int.sizeof;
                case VARENUM.VT_VARIANT:
                    return (VARIANT.sizeof + 3) / int.sizeof;
                default:
            }

            return 1;
        }

        // Like DispCallFunc, but using delegates

        size_t paramCount;
        for (int i = 0; i < paramTypes.length; i++) {
            paramCount += variantSize(paramTypes[i]);
        }

        auto argptr = cast(int*)HeapAlloc(GetProcessHeap(), 0, paramCount * int.sizeof);

        uint pos;
        for (int i = 0; i < paramTypes.length; i++) {
            VARIANT* p = args[i];
            if (paramTypes[i] == VARENUM.VT_VARIANT)
                memcpy(&argptr[pos], p, variantSize(paramTypes[i]) * int.sizeof);
            else
                memcpy(&argptr[pos], &p.lVal, variantSize(paramTypes[i]) * int.sizeof);
            pos += variantSize(paramTypes[i]);
        }

        int ret = 0;

        switch (paramCount) {
            case 0: ret = method(); break;
            case 1: ret = (cast(int delegate(int))method)(argptr[0]); break;
            case 2: ret = (cast(int delegate(int, int))method)(argptr[0], argptr[1]); break;
            case 3: ret = (cast(int delegate(int, int, int))method)(argptr[0], argptr[1], argptr[2]); break;
            case 4: ret = (cast(int delegate(int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3]); break;
            case 5: ret = (cast(int delegate(int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4]); break;
            case 6: ret = (cast(int delegate(int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5]); break;
            case 7: ret = (cast(int delegate(int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6]); break;
            case 8: ret = (cast(int delegate(int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7]); break;
            case 9: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8]); break;
            case 10: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9]); break;
            case 11: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10]); break;
            case 12: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11]); break;
            case 13: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12]); break;
            case 14: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13]); break;
            case 15: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14]); break;
            case 16: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14], argptr[15]); break;
            case 17: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14], argptr[15], argptr[16]); break;
            case 18: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14], argptr[15], argptr[16], argptr[17]); break;
            case 19: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14], argptr[15], argptr[16], argptr[17], argptr[18]); break;
            case 20: ret = (cast(int delegate(int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int))method)(argptr[0], argptr[1], argptr[2], argptr[3], argptr[4], argptr[5], argptr[6], argptr[7], argptr[8], argptr[9], argptr[10], argptr[11], argptr[12], argptr[13], argptr[14], argptr[15], argptr[16], argptr[17], argptr[18], argptr[19]); break;
            default:
                return DISP_E_BADPARAMCOUNT;
        }

        if (result !is null && returnType != VARENUM.VT_VOID) {
            result.vt = returnType;
            result.lVal = ret;
        }

        HeapFree(GetProcessHeap(), 0, argptr);
        return S_OK;
    }

}

/**
*/
class EventProvider(T) : Implements!(T) {

    extern(D):

    private MethodProxy[int] methodTable_;
    private int[string] nameTable_;

    private IConnectionPoint connectionPoint_;
    private uint cookie_;

    /**
    */
    this(IUnknown source) {
        auto cpc = ComPtr!IConnectionPointContainer(source);
        assert(!cpc.isNull);

        if (cpc.FindConnectionPoint(&uuidof!(T), &connectionPoint_) != S_OK)
            throw new ArgumentException("Source object does not expose '" ~ T.stringof ~ "' event interface.");

        if (connectionPoint_.Advise(this, &cookie_) != S_OK) {
            cookie_ = 0;
            tryRelease(connectionPoint_);
            throw new InvalidOperationException("Could not Advise() the event interface '" ~ T.stringof ~ "'.");
        }

        if (connectionPoint_ is null || cookie_ == 0) {
            if (connectionPoint_ !is null)
                tryRelease(connectionPoint_);
            throw new ArgumentException("Connection point for event interface '" ~ T.stringof ~ "' cannot be created.");
        }
    }

    /*
    ~this() {
    //disconnect();
    }
    */

    void disconnect() {
        if (connectionPoint_ !is null && cookie_ != 0) {
            try {
                connectionPoint_.Unadvise(cookie_);
            }
            finally {
                tryRelease(connectionPoint_);
                connectionPoint_ = null;
                cookie_ = 0;
            }
        }
    }

    /**
    */
    void bind(ID, R, P...)(ID member, R delegate(P) handler) {
        static if (is(ID : string)) {
            bool found;
            int dispId = DISPID_UNKNOWN;
            if (tryFindDispId(member, dispId))
                bind(dispId, handler);
            else
                throw new ArgumentException("Member '" ~ member ~ "' not found in type '" ~ T.stringof ~ "'.");
        }
        else static if (is(ID : int)) {
            MethodProxy m = handler;
            methodTable_[member] = m;
        }
    }

    private bool tryFindDispId(string name, out int dispId) {

        void ensureNameTable() {
            if (nameTable_ == null) {
                scope clsidKey = RegistryKey.classesRoot.openSubKey("Interface\\" ~ uuidof!(T).toString());
                if (clsidKey !is null) {
                    scope typeLibRefKey = clsidKey.openSubKey("TypeLib");
                    if (typeLibRefKey !is null) {
                        string typeLibVersion = typeLibRefKey.getValue!(string)("Version");
                        if (typeLibVersion == null) {
                            scope versionKey = clsidKey.openSubKey("Version");
                            if (versionKey !is null)
                                typeLibVersion = versionKey.getValue!(string)(null);
                        }

                        scope typeLibKey = RegistryKey.classesRoot.openSubKey("TypeLib\\" ~ typeLibRefKey.getValue!(string)(null));
                        if (typeLibKey !is null) {
                            scope pathKey = typeLibKey.openSubKey(typeLibVersion ~ "\\0\\Win32");
                            if (pathKey !is null) {
                                ITypeLib typeLib;
                                if (LoadTypeLib(pathKey.getValue!(string)(null).toUTF16z(), &typeLib) == S_OK) {
                                    scope(exit) tryRelease(typeLib);

                                    ITypeInfo typeInfo;
                                    if (typeLib.GetTypeInfoOfGuid(&uuidof!(T), typeInfo) == S_OK) {
                                        scope(exit) tryRelease(typeInfo);

                                        TYPEATTR* typeAttr;
                                        if (typeInfo.GetTypeAttr(typeAttr) == S_OK) {
                                            scope(exit) typeInfo.ReleaseTypeAttr(typeAttr);

                                            for (uint i = 0; i < typeAttr.cFuncs; i++) {
                                                FUNCDESC* funcDesc;
                                                if (typeInfo.GetFuncDesc(i, funcDesc) == S_OK) {
                                                    scope(exit) typeInfo.ReleaseFuncDesc(funcDesc);

                                                    wchar* bstrName;
                                                    if (typeInfo.GetDocumentation(funcDesc.memid, &bstrName, null, null, null) == S_OK) {
                                                        string memberName = fromBstr(bstrName);
                                                        nameTable_[memberName.toLower()] = funcDesc.memid;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        dispId = DISPID_UNKNOWN;

        ensureNameTable();

        if (auto value = name.toLower() in nameTable_) {
            dispId = *value;
            return true;
        }

        return false;
    }

    extern(Windows):

    override int Invoke(int dispIdMember, REFGUID riid, uint lcid, ushort wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, uint* puArgError) {
        if (*riid != emptyGUID)
            return DISP_E_UNKNOWNINTERFACE;

        try {
            if (auto handler = dispIdMember in methodTable_) {
                VARIANT*[8] args;
                for (int i = 0; i < handler.paramTypes.length && i < 8; i++) {
                    args[i] = &pDispParams.rgvarg[handler.paramTypes.length - i - 1];
                }

                VARIANT result;
                if (pVarResult == null)
                    pVarResult = &result;

                int hr = handler.invoke(args, pVarResult);

                for (int i = 0; i < handler.paramTypes.length; i++) {
                    if (args[i].vt == (VT_BYREF | VT_BOOL)) {
                        // Fix bools to VARIANT_BOOL
                        *args[i].pboolVal = (*args[i].pboolVal == 0) ? VARIANT_FALSE : VARIANT_TRUE;
                    }
                }

                return hr;
            }
            else
                return DISP_E_MEMBERNOTFOUND;
        }
        catch {
            return E_FAIL;
        }

        return S_OK;
    }

}

// Forces compilation of template class
unittest {
    interface TestEvents : IDispatch {
        mixin(uuid("00000000-272f-11d2-836f-0000f87a7782"));
    }
    EventProvider!(TestEvents) events;
}

