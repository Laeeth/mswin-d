/**
* Provides support for COM (Component Object Model).
*
* See $(LINK2 http://msdn.microsoft.com/en-us/library/ms690233(VS.85).aspx, MSDN) for a glossary of terms.
*
* Copyright: (c) 2009 John Chapman
*
* License: See $(LINK2 ..\..\..\licence.txt, licence.txt) for use and distribution terms.
*/
module mswin.com;

import 
std.string,
    std.stream,
    std.typetuple,
    std.traits,
    std.exception,
    std.conv;
import std.algorithm;
import std.array;
import std.utf : toUTF8, toUTF16z;
import core.exception, core.memory;

import std.system;
import std.uuid;

public import win32.windows, win32.unknwn, win32.objbase, win32.objidl, win32.oaidl, win32.oleauto, win32.wtypes;

public import mswin.exception;

debug import std.stdio : writefln;

pragma(lib, "ole32.lib");

//////////////////////////////////////////////////////////////////////////////////////////
// COM init/shutdown
//////////////////////////////////////////////////////////////////////////////////////////

package bool isComAlive = false;

public void comInit()
{
    if (isComAlive == false)
    {
        CoInitializeEx(null, COINIT.COINIT_APARTMENTTHREADED).checkResult();
        isComAlive = true;
    }
}

public void comShutdown()
{
    // Before we shut down COM, give classes a chance to release any COM resources.  
    try {
        GC.collect();
    }
    finally
    {
        if (isComAlive)
        {
            isComAlive = false;
            CoUninitialize();
        }
    }
}

/*static this()
{
    comInit();
}

static ~this()
{
    comShutdown();
}*/

/**
* Encapsulate a COM interface pointer, to provide 
* automatic reference count management. 
* Examples:
* ---
*  ComPtr!IShellFolder desktopFolder;
*  SHGetDesktopFolder(desktopFolder.pptr).checkResult();
*  ComPtr!IEnumIDList enumID;
*  desktopFolder.EnumObjects(null, SHCONTF.SHCONTF_FOLDERS, enumID.pptr).checkResult();
*/
struct ComPtr(T) if (is(T : IUnknown)) 
{
    T _obj;
    alias _obj this;

    this(T obj)
    {
        _obj = obj;
        _addRef();
    }

    static if(!is(T == IUnknown))
    this(IUnknown obj) 
    {
        if(obj !is null)
            obj.QueryInterface(cast(GUID*)&uuidof!(T), retval(_obj)).checkResult();
    }

    /**
    * Creates an object of the class associated with a specified GUID.
    * Params:
    *   clsid = The class associated with the object.
    *   context = Context in which the code that manages the object will run.
    *   outer = If null, indicates that the object is not being created as part of an aggregate.
    */
    this(GUID clsid, ExecutionContext context = ExecutionContext.All, IUnknown outer=null)
    {
        if(context & (CLSCTX.CLSCTX_LOCAL_SERVER | CLSCTX.CLSCTX_REMOTE_SERVER)) {
            ComPtr!IUnknown unk;
            CoCreateInstance(&clsid, outer, context, &uuidof!(IUnknown), retval(unk)).checkResult();
            OleRun(unk).checkResult();
            unk.QueryInterface(cast(GUID*)&uuidof!(T), retval(_obj)).checkResult();
        }
        else 
            CoCreateInstance(&clsid, outer, context, &uuidof!(T), retval(_obj)).checkResult();
    }

    /**
    * Creates an object of the class associated with a specified progid.
    * Params:
    *   progid = The progid of the class associated with the object.
    *   context = Context in which the code that manages the object will run.
    *   outer = If null, indicates that the object is not being created as part of an aggregate.
    */
    this(string progid, ExecutionContext context = ExecutionContext.All, IUnknown outer=null)
    {
        GUID clsid;
        CLSIDFromProgID(std.utf.toUTF16z(progid), &clsid).checkResult();
        this(clsid, context, outer);
    }

    this(this)
    {
        _addRef();
    }

    ~this() 
    {
        _release();
    }

    void opAssign(ComPtr rhs)
    {
        swap(this, rhs);
    }

    void opAssign(IUnknown obj)
    {
        T oldObj = _obj;
        _obj = null;
        scope(exit) {
            if(oldObj !is null)
                oldObj.Release();
        }

        obj.QueryInterface(cast(GUID*)&uuidof!(T), retval(_obj)).checkResult();
    }

    static ComPtr opCall(U)(U obj) 
    {
        return ComPtr(obj);
    }

    bool isNull() 
    {
        return (_obj is null);
    }

    auto opCast(C)()
    {  
        static if(is(C == bool))
        {
            return _obj !is null;
        }
        else static if(is(C == T))
        {
            return _obj;
        }
        else static assert(false);
    }

    /// Returns the address of the interface pointer contained in this
    /// class. This is useful when using the COM/OLE interfaces to create
    /// this interface. The previous hold interface is Release()'ed
    @property T* pptr()
    {
        _release();
        _obj = null;
        return &_obj;
    }

    /// Saves/sets the interface only AddRef()ing if incrementRefCount is true.
    /// This call will release any previously acquired interface.
    void attach(T obj, bool incrementRefCount=false)
    {
        _release();
        _obj = obj;
        if(incrementRefCount)
            _addRef();
    }

    /// Simply null-ify the interface pointer so that it isn't Released()'ed.
    T detach()
    {
        T old = _obj;
        _obj = null;
        return old;
    }


private:


    void _addRef() 
    {
        if (_obj !is null)
            _obj.AddRef();
    }

    /// Release the hold interface if is not null.
    /// The interface pointer is NOT set to null afterwards.
    void _release() 
    {
        if (_obj !is null) 
            _obj.Release();
    }

}

unittest {
    assert(comInit());
    scope(exit) comShutdown();
}


//////////////////////////////////////////////////////////////////////////////////////////
// GUID generation/manipulation
//////////////////////////////////////////////////////////////////////////////////////////

static GUID emptyGUID = { 0x0, 0x0, 0x0, [0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0] };

/**
* Initializes a new GUID instance using the specified integers and bytes.
* Params:
*   a = The first 4 bytes.
*   b = The next 2 bytes.
*   c = The next 2 bytes.
*   d = The next byte.
*   e = The next byte.
*   f = The next byte.
*   g = The next byte.
*   h = The next byte.
*   i = The next byte.
*   j = The next byte.
*   k = The next byte.
* Returns: The resulting GUID.
*/
GUID guid(uint a, ushort b, ushort c, ubyte d, ubyte e, ubyte f, ubyte g, ubyte h, ubyte i, ubyte j, ubyte k) {
    return GUID(a, b, c, [d, e, f, g, h, i, j, k]);
}

/**
* Initializes _a new instance using the specified integers and byte array.
* Params:
*   a = The first 4 bytes.
*   b = The next 2 bytes.
*   c = The next 2 bytes.
*   d = The remaining 8 bytes.
* Returns: The resulting GUID.
* Throws: IllegalArgumentException if d is not 8 bytes long.
*/
GUID guid(uint a, ushort b, ushort c, ubyte[] d) {
    if (d.length != 8)
        throw new ArgumentException("Byte array for GUID must be 8 bytes long.");
    return guid(a, b, c, d[0..8]);
}

/**
* Initializes a new instance using the value represented by the specified string.
* Params: s = A string containing a GUID in groups of 8, 4, 4, 4 and 12 digits with hyphens between the groups. The GUID can optionally be enclosed in braces.
* Returns: The resulting GUID.
*/
GUID guid(string s) {

    ulong parse(string s) {

        bool hexToInt(char c, out uint result) {
            if (c >= '0' && c <= '9') result = c - '0';
            else if (c >= 'A' && c <= 'F') result = c - 'A' + 10;
            else if (c >= 'a' && c <= 'f') result = c - 'a' + 10;
            else result = -1;
            return (cast(int)result >= 0);
        }

        ulong result;
        uint value, index;
        while (index < s.length && hexToInt(s[index], value)) {
            result = result * 16 + value;
            index++;
        }
        return result;
    }

    s = s.strip();

    if (s[0] == '{') {
        s = s[1 .. $];
        if (s[$ - 1] == '}')
            s = s[0 .. $ - 1];
    }

    if (s[0] == '[') {
        s = s[1 .. $];
        if (s[$ - 1] == ']')
            s = s[0 .. $ - 1];
    }

    if (s.find("-").empty)
        throw new FormatException("Unrecognised GUID format.");

    GUID self;
    self.Data1 = cast(uint)parse(s[0 .. 8]);
    self.Data2 = cast(ushort)parse(s[9 .. 13]);
    self.Data3 = cast(ushort)parse(s[14 .. 18]);
    uint m = cast(uint)parse(s[19 .. 23]);
    self.Data4[0] = cast(ubyte)(m >> 8);
    self.Data4[1] = cast(ubyte)m;
    ulong n = parse(s[24 .. $]);
    m = cast(uint)(n >> 32);
    self.Data4[2] = cast(ubyte)(m >> 8);
    self.Data4[3] = cast(ubyte)m;
    m = cast(uint)n;
    self.Data4[4] = cast(ubyte)(m >> 24);
    self.Data4[5] = cast(ubyte)(m >> 16);
    self.Data4[6] = cast(ubyte)(m >> 8);
    self.Data4[7] = cast(ubyte)m;
    return self;
}

/**
* Creates a new GUID.
*/
GUID guid() {
    GUID uid;

    CoCreateGuid(&uid).checkResult();

    return uid;
}

/// ditto
string toString(const GUID uid) 
{
    void hexToString(ref char[] s, ref uint index, uint a, uint b) {

        char hexToChar(uint a) {
            a = a & 0x0F;
            return cast(char)((a > 9) ? a - 10 + 0x61 : a + 0x30);
        }

        s[index++] = hexToChar(a >> 4);
        s[index++] = hexToChar(a);
        s[index++] = hexToChar(b >> 4);
        s[index++] = hexToChar(b);
    }

    char[] s = new char[38];
    uint index = 0;
    s[index++] = '{';
    s[$ - 1] = '}';

    hexToString(s, index, uid.Data1 >> 24, uid.Data1 >> 16);
    hexToString(s, index, uid.Data1 >> 8, uid.Data1);
    s[index++] = '-';
    hexToString(s, index, uid.Data2 >> 8, uid.Data2);
    s[index++] = '-';
    hexToString(s, index, uid.Data3 >> 8, uid.Data3);
    s[index++] = '-';
    hexToString(s, index, uid.Data4[0], uid.Data4[1]);
    s[index++] = '-';
    hexToString(s, index, uid.Data4[2], uid.Data4[3]);
    hexToString(s, index, uid.Data4[4], uid.Data4[5]);
    hexToString(s, index, uid.Data4[6], uid.Data4[7]);

    return cast(string)s;
}

/**
* Translates a std.uuid to a GUID
*/
GUID guid(std.uuid.UUID uid) {
    GUID ans;
    // Easy when slice-able
    auto quickFill = (cast(ubyte*)&ans)[0..16];
    quickFill[] = uid.data[];

    // GUID Uses Native Endian in first three groupings
    // UUID is all big Endian
    if(endian == Endian.littleEndian) {
        quickFill[0..4].reverse();
        quickFill[4..6].reverse();
        quickFill[6..8].reverse();
    }
    // Last 8 bytes are the same

    return ans;
}

/**
* Retrieves the ProgID for a given class identifier (CLSID).
*/
string progIdFromClsid(GUID clsid) {
    wchar* str;
    ProgIDFromCLSID(&clsid, &str).checkResult();
    scope(exit) CoTaskMemFree(str);
    return toUTF8(str[0 .. wcslen(str)]);
}

/**
* Retrieves the class identifier (CLSID) for a given ProgID.
*/
GUID clsidFromProgId(string progId) {
    GUID clsid;
    CLSIDFromProgID(progId.toUTF16z(), &clsid).checkResult();
    return clsid;
}

unittest {
    import std.typecons;
    Tuple!(GUID, GUID) getBoth(string g) {
        return tuple(guid(UUID(g)), guid(g));
    }
    auto ids = getBoth("2933bf8f-7b36-11d2-b20e-00c04f983e60");
    assert(ids[0] == ids[1]);

    ids = getBoth("2933bf80-7b36-11d2-b20e-00c04f983e60");
    assert(ids[0] == ids[1]);

    ids = getBoth("8c033caa-6cd6-4f73-b728-4531af74945f");
    assert(ids[0] == ids[1]);

    ids = getBoth("0c05d096-f45b-4aca-ad1a-aa0bc25518dc");
    assert(ids[0] == ids[1]);
}

/**
* Translates a GUID to a std.uuid
*/
std.uuid.UUID uuid(GUID uid) {
    std.uuid.UUID ans;
    // Easy when slice-able
    ans.data[] = (cast(ubyte*)&uid)[0..16];

    // GUID Uses Native Endian in first three groupings
    // UUID is all big Endian
    if(endian == Endian.littleEndian) {
        ans.data[0..4].reverse();
        ans.data[4..6].reverse();
        ans.data[6..8].reverse();
    }
    // Last 8 bytes are the same

    return ans;
}

unittest {
    import std.typecons;
    Tuple!(UUID, UUID) getBoth(string g) {
        return tuple(uuid(guid(g)), UUID(g));
    }
    auto ids = getBoth("2933bf8f-7b36-11d2-b20e-00c04f983e60");
    assert(ids[0] == ids[1]);

    ids = getBoth("2933bf80-7b36-11d2-b20e-00c04f983e60");
    assert(ids[0] == ids[1]);

    ids = getBoth("8c033caa-6cd6-4f73-b728-4531af74945f");
    assert(ids[0] == ids[1]);

    ids = getBoth("0c05d096-f45b-4aca-ad1a-aa0bc25518dc");
    assert(ids[0] == ids[1]);
}

/**
* Associates a GUID with an interface.
* Params: g = A string representing the GUID in normal registry format with or without the { } delimiters.
* Examples:
* ---
* interface IXMLDOMDocument2 : IDispatch {
*   mixin(uuid("2933bf95-7b36-11d2-b20e-00c04f983e60"));
* }
*
* // Expands to the following code:
* //
* // interface IXMLDOMDocument2 : IDispatch {
* //   static GUID IID = { 0x2933bf95, 0x7b36, 0x11d2, [0xb2, 0x0e, 0x00, 0xc0, 0x4f, 0x98, 0x3e, 0x60] };
* // }
* ---
*/
string uuid(string g) {
    if (g.length == 38) {
        assert(g[0] == '{' && g[$-1] == '}', "Incorrect format for GUID.");
        return uuid(g[1..$-1]);
    }
    else if (g.length == 36) {
        assert(g[8] == '-' && g[13] == '-' && g[18] == '-' && g[23] == '-', "Incorrect format for GUID.");
        return "static const GUID IID = { 0x" ~ g[0..8] ~ ",0x" ~ g[9..13] ~ ",0x" ~ g[14..18] ~ ", [0x" ~ g[19..21] ~ ",0x" ~ g[21..23] ~ ",0x" ~ g[24..26] ~ ",0x" ~ g[26..28] ~ ",0x" ~ g[28..30] ~ ",0x" ~ g[30..32] ~ ",0x" ~ g[32..34] ~ ",0x" ~ g[34..36] ~ "] };";
    }
    else assert(false, "Incorrect format for GUID.");
}

/**
* Retrieves the GUID associated with the specified variable or type.
* Examples:
* ---
* import mswin.com,
*   std.stdio;
*
* void main() {
*   writefln("The GUID of IXMLDOMDocument2 is %s", uuidof!(IXMLDOMDocument2));
* }
*
* // Produces:
* // The GUID of IXMLDOMDocument2 is {2933bf95-7b36-11d2-b20e-00c04f983e60}
* ---
*/
template uuidof(alias T) {
    static if (is(typeof(T)))
        const GUID uuidof = uuidofT!(typeof(T));
    else
        const GUID uuidof = uuidofT!(T);
}

template uuidofT(T : T) {
    static if (is(typeof(mixin("IID_" ~ T.stringof))))
        const GUID uuidofT = mixin("IID_" ~ T.stringof); // e.g., IID_IShellFolder
    else static if (is(typeof(mixin("CLSID_" ~ T.stringof))))
        const GUID uuidofT = mixin("CLSID_" ~ T.stringof); // e.g., CLSID_Shell
    else static if (is(typeof(T.IID)))
        const GUID uuidofT = T.IID;
    else
        static assert(false, "No GUID has been associated with '" ~ T.stringof ~ "'.");
}

//////////////////////////////////////////////////////////////////////////////////////////
// Common interfaces implementation
//////////////////////////////////////////////////////////////////////////////////////////

/**
* Decrements the reference count for an object.
*/
void tryRelease(IUnknown obj) {
    if (obj) {
        try {
            obj.Release();
        }
        catch {
        }
    }
}

void** retval(T)(out T ppv)
in {
    assert(&ppv != null);
}
body {
    return cast(void**)&ppv;
}

/// Specifies the context in which the code that manages an object will run.
/// See_Also: $(LINK2 http://msdn.microsoft.com/en-us/library/ms693716.aspx, CLSCTX Enumeration).
enum ExecutionContext : uint {
    InProcessServer  = CLSCTX.CLSCTX_INPROC_SERVER,  /// The code that creates and manages objects of this class is a DLL that runs in the same process as the caller of the function specifying the class context.
    InProcessHandler = CLSCTX.CLSCTX_INPROC_HANDLER, /// The code that manages objects of this class is an in-process handler. is a DLL that runs in the client process and implements client-side structures of this class when instances of the class are accessed remotely.
    LocalServer      = CLSCTX.CLSCTX_LOCAL_SERVER,   /// The code that creates and manages objects of this class runs on same machine but is loaded in a separate process space.
    RemoteServer     = CLSCTX.CLSCTX_REMOTE_SERVER,  /// A  remote context. The code that creates and manages objects of this class is run on a different computer.
    All              = CLSCTX_ALL
}


template Interfaces(TList...) {

    static ComPtr!T coCreate(T)(ExecutionContext context = ExecutionContext.All) {
        import std.typetuple;
        static if (std.typetuple.IndexOf!(T, TList) == -1)
            static assert(false, "'" ~ typeof(this).stringof ~ "' does not support '" ~ T.stringof ~ "'.");
        else
            return ComPtr!T(uuidof!(typeof(this)), context);
    }
}

template QueryInterfaceImpl(TList...) {

    extern(Windows):
    int QueryInterface(GUID* riid, void** ppvObject) {
        if (ppvObject is null || riid is null)
            return E_POINTER;

        *ppvObject = null;

        if (*riid == uuidof!(IUnknown)) {
            *ppvObject = cast(void*)cast(IUnknown)this;
        }
        else foreach (T; TList) {
            // Search the specified list of types to see if we support the interface we're being asked for.
            if (*riid == uuidof!(T)) {
                // This is the one, so we need look no further.
                *ppvObject = cast(void*)cast(T)this;
                break;
            }
        }

        if (*ppvObject is null)
            return E_NOINTERFACE;

        (cast(IUnknown)this).AddRef();
        return S_OK;
    }

}

// Implements AddRef & Release for IUnknown subclasses.
template ReferenceCountImpl() {

    private int refCount_ = 1;
    private bool finalized_;

    extern(Windows):

    uint AddRef() {
        return ++refCount_;
    }

    uint Release() {
        if (--refCount_ == 0) {
            if (!finalized_) {
                finalized_ = true;
                runFinalizer(this);
            }

            core.memory.GC.removeRange(cast(void*)this);
            core.memory.GC.free(cast(void*)this);
        }
        return refCount_;
    }

    extern(D):

    // IUnknown subclasses must manage their memory manually.
    new(size_t sz) {
        void* p = std.c.stdlib.malloc(sz);
        if (p is null)
            throw new OutOfMemoryError;

        core.memory.GC.addRange(p, sz);

        return p;
    }

}

template InterfacesTuple(T) {

    static if (is(T == Object)) {
        alias TypeTuple!() InterfacesTuple;
    }
    static if (is(BaseTypeTuple!(T)[0] == Object)) {
        alias TypeTuple!(BaseTypeTuple!(T)[1 .. $]) InterfacesTuple;
    }
    else {
        alias std.typetuple.NoDuplicates!(
                                          TypeTuple!(BaseTypeTuple!(T)[1 .. $],
                                                     InterfacesTuple!(BaseTypeTuple!(T)[0])))
            InterfacesTuple;
    }

}

/// Provides an implementation of IUnknown suitable for using as mixin.
template IUnknownImpl(T...) {

    static if (is(T[0] : Object))
        mixin QueryInterfaceImpl!(InterfacesTuple!(T[0]), T[1 .. $]);
    else
        mixin QueryInterfaceImpl!(T);
    mixin ReferenceCountImpl;

}

/// Provides an implementation of IDispatch suitable for using as mixin.
template IDispatchImpl(T...) {

    mixin IUnknownImpl!(T);

    int GetTypeInfoCount(uint* pctinfo) {
        assert(pctinfo);
        *pctinfo = 0;
        return E_NOTIMPL;
    }

    int GetTypeInfo(uint iTInfo, uint lcid, ITypeInfo* ppTInfo) {
        assert(ppTInfo);
        *ppTInfo = null;
        return E_NOTIMPL;
    }

    int GetIDsOfNames(REFGUID riid, wchar** rgszNames, uint cNames, uint lcid, int* rgDispId) {
        return E_NOTIMPL;
    }

    int Invoke(int dispIdMember, REFGUID riid, uint lcid, ushort wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, uint* puArgError) {
        return DISP_E_UNKNOWNNAME;
    }

}

template AllBaseTypesOfImpl(T...) {

    static if (T.length == 0)
        alias TypeTuple!() AllBaseTypesOfImpl;
    else
        alias TypeTuple!(T[0],
                         AllBaseTypesOfImpl!(std.traits.BaseTypeTuple!(T[0])),
                         AllBaseTypesOfImpl!(T[1 .. $]))
            AllBaseTypesOfImpl;

}

template AllBaseTypesOf(T...) {

    alias NoDuplicates!(AllBaseTypesOfImpl!(T)) AllBaseTypesOf;

}

/**
* The abstract base class for COM objects that derive from IUnknown or IDispatch.
*
* The Implements class provides default implementations of methods required by those interfaces. Therefore, subclasses need only override them when they
* specifically need to provide extra functionality. This class also overrides the new operator so that instances are not garbage collected.
* Examples:
* ---
* class MyImpl : Implements!(IUnknown) {
* }
* ---
*/
abstract class Implements(T...) : T {

    static if (IndexOf!(IDispatch, AllBaseTypesOf!(T)) != -1)
        mixin IDispatchImpl!(T, AllBaseTypesOf!(T));
    else
        mixin IUnknownImpl!(T, AllBaseTypesOf!(T));

}

// DMD prevents destructors from running on COM objects.
void runFinalizer(Object obj) {
    if (obj) {
        ClassInfo** ci = cast(ClassInfo**)cast(void*)obj;
        if (*ci) {
            if (auto c = **ci) {
                do {
                    if (c.destructor) {
                        auto finalizer = cast(void function(Object))c.destructor;
                        finalizer(obj);
                    }
                    c = c.base;
                } while (c);
            }
        }
    }
}

/**
* Indicates whether the specified object represents a COM object.
* Params: obj = The object to check.
* Returns: true if obj is a COM type; otherwise, false.
*/
bool isComObject(Object obj) {
    ClassInfo** ci = cast(ClassInfo**)cast(void*)obj;
    if (*ci !is null) {
        ClassInfo c = **ci;
        if (c !is null)
            return ((c.flags & 1) != 0);
    }
    return false;
}

//////////////////////////////////////////////////////////////////////////////////////////
// COM server helpers
//////////////////////////////////////////////////////////////////////////////////////////

/**
* Contains boiler-plate code for creating a COM _server (a DLL that exports COM classes).
* Examples:
* ---
* --- hello.d ---
* module hello;
*
* // This is the interface.
*
* private import mswin.com;
*
* interface ISaysHello : IUnknown {
*   mixin(uuid("ae0dd4b7-e817-44ff-9e11-d1cffae11f16"));
*
*   int sayHello();
* }
*
* // coclass
* abstract class SaysHello {
*   mixin(uuid("35115e92-33f5-4e14-9d0a-bd43c80a75af"));
*
*   mixin Interfaces!(ISaysHello);
* }
* ---
*
* ---
* --- server.d ---
* module server;
*
* // This is the DLL's private implementation.
*
* import mswin.com, mswin.registry, hello;
*
* mixin Export!(SaysHelloClass);
*
* // Implements ISaysHello
* class SaysHelloClass : Implements!(ISaysHello) {
*   // Note: must have the same CLSID as the SaysHello coclass above.
*   mixin(uuid("35115e92-33f5-4e14-9d0a-bd43c80a75af"));
*
*   int sayHello() {
*     writefln("Hello there!");
*     return S_OK;
*   }
*
* }
* ---
*
* ---
* --- client.d ---
* module client;
*
* import mswin.com, hello;
*
* void main() {
*   ISaysHello saysHello = SaysHello.comCreate!(ISaysHello);
*   saysHello.sayHello(); // Prints "Hello there!"
*   saysHello.Release();
* }
* ---
*
* The COM _server needs to be registered with the system. Usually, a CLSID is associated with the DLL in the registry 
* (under HKEY_CLASSES_ROOT\CLSID). On Windows XP and above, an alternative is to deploy an application manifest in the same folder 
* as the client application. This is an XML file that does the same thing as the registry method. Here's an example:
*
* ---
* --- client.exe.manifest ---
* <?xml version="1.0" encoding="utf-8" standalone="yes"?>
* <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
*   <assemblyIdentity name="client.exe" version="1.0.0.0" type="win32"/>
*   <file name="C:\\Program Files\\My COM Server\\server.dll">
*     <comClass clsid="{35115e92-33f5-4e14-9d0a-bd43c80a75af}" description="SaysHello" threadingModel="Apartment"/>
*  </file>
* </assembly>
*
* ---
* Alternatively, define a static register and unregister method on each coclass implementation. If the methods exist, the DLL will 
* register itself in the registry when 'regsvr32' is executed, and unregister itself on 'regsvr32 /u'.
*
*/

import core.runtime, core.sys.windows.dll;

alias void* Handle;

private __gshared Handle moduleHandle_;
private __gshared int serverLockCount_;

///
Handle getHInstance() {
    return moduleHandle_;
}

///
void setHInstance( Handle value) {
    moduleHandle_ = value;
}

///
string getLocation()
{
    wchar[MAX_PATH] buffer;
    uint len = GetModuleFileName(moduleHandle_, buffer.ptr, buffer.length);
    if(!len)
        throw new Win32Exception();
    else if(len == MAX_PATH) {
        if(GetLastError() == ERROR_INSUFFICIENT_BUFFER)
            throw new Win32Exception();
    }
    return to!string(buffer[0..len]);
}

unittest {
    assert(getLocation().length > 0);
}

///
int getServerLockCount() {
    return serverLockCount_;
}

///
void lockServer() {
    InterlockedIncrement(serverLockCount_);
}

///
void unlockServer() {
    InterlockedDecrement(serverLockCount_);
}

///
class ClassFactory(T) : Implements!(IClassFactory) {

    int CreateInstance(IUnknown pUnkOuter, const ref GUID riid, void** ppvObject) {
        if (pUnkOuter !is null && riid != uuidof!(IUnknown))
            return CLASS_E_NOAGGREGATION;

        *ppvObject = null;
        int hr = E_OUTOFMEMORY;

        T obj = new T;
        if (obj !is null) {
            hr = obj.QueryInterface(riid, ppvObject);
            obj.Release();
        }
        return hr;
    }

    int LockServer(int fLock) {
        if (fLock)
            lockServer();
        else
            unlockServer();
        return S_OK;
    }

}

bool registerComClass(ComClass)() {
    bool success;

    try {
        scope clsidKey = RegistryKey.classesRoot.createSubKey("CLSID\\" ~ uuidof!(ComClass).toString());
        if (clsidKey !is null) {
            clsidKey.setValue!(string)(null, ComClass.classinfo.name ~ " Class");

            scope subKey = clsidKey.createSubKey("InprocServer32");
            if (subKey !is null) {
                subKey.setValue!(string)(null, getLocation());
                subKey.setValue!(string)("ThreadingModel", "Apartment");

                scope progIDSubKey = clsidKey.createSubKey("ProgID");
                if (progIDSubKey !is null) {
                    progIDSubKey.setValue!(string)(null, ComClass.classinfo.name);

                    scope progIDKey = RegistryKey.classesRoot.createSubKey(ComClass.classinfo.name);
                    if (progIDKey !is null) {
                        progIDKey.setValue!(string)(null, ComClass.classinfo.name ~ " Class");

                        scope clsidSubKey = progIDKey.createSubKey("CLSID");
                        if (clsidSubKey !is null)
                            clsidSubKey.setValue!(string)(null, uuidof!(ComClass).toString());
                    }
                }
            }
        }

        success = true;
    }
    catch {
        success = false;
    }

    return success;
}

bool unregisterComClass(ComClass)() {
    bool success;

    try {
        scope clsidKey = RegistryKey.classesRoot.openSubKey("CLSID", true);
        if (clsidKey !is null)
            clsidKey.deleteSubKeyTree(uuidof!(ComClass).toString());

        RegistryKey.classesRoot.deleteSubKeyTree(ComClass.classinfo.name);

        success = true;
    }
    catch {
        success = false;
    }

    return success;
}


int DllMainImpl(Handle hInstance, uint dwReason, void* pvReserved)
{
    if (dwReason == 1 /*DLL_PROCESS_ATTACH*/) {

        dll_process_attach( hInstance, false );
        setHInstance(hInstance);

        return 1;
    }
    else if (dwReason == 0 /*DLL_PROCESS_DETACH*/) {

        dll_process_detach( hInstance, true );

        return 1;
    }
    else if (dwReason == 2 /*DLL_THREAD_ATTACH*/) {

        dll_thread_attach( true, true );

        return 1;
    }
    else if (dwReason == 3 /*DLL_THREAD_DETACH*/) {

        dll_thread_detach( true, true );

        return 1;
    }

    return 0;
}

int DllCanUnloadNowImpl()
{
    int i = getServerLockCount(); 
    return (i == 0) ? S_OK : S_FALSE;
}

///
mixin template Export(T...) {

    import mswin.com;

    extern(Windows) int DllMain(Handle hInstance, uint dwReason, void* pvReserved)
    {
        return DllMainImpl(hInstance, dwReason, pvReserved);
    }


    extern(Windows)
        int DllGetClassObject(ref GUID rclsid, ref GUID riid, void** ppv) {
            int hr = CLASS_E_CLASSNOTAVAILABLE;
            *ppv = null;

            foreach (coclass; T) {
                if (rclsid == uuidof!(coclass)) {
                    IClassFactory factory = new ClassFactory!(coclass);
                    if (factory is null)
                        return E_OUTOFMEMORY;
                    scope(exit) tryRelease(factory);

                    hr = factory.QueryInterface(riid, ppv);
                }
            }

            return hr;
        }

    extern(Windows)
        int DllCanUnloadNow() {
            return DllCanUnloadNowImpl();
        }

    extern(Windows) int DllRegisterServer() {
        bool success;

        foreach (coclass; T) {

            static if (is(typeof(coclass.register))) {
                static assert(is(typeof(coclass.unregister)), "'register' must be matched by a corresponding 'unregister' in '" ~ coclass.stringof ~ "'.");

                success = registerComClass!(coclass)();
                coclass.register();
            }
        }

        return success ? S_OK : SELFREG_E_CLASS;
    }

    extern(Windows) int DllUnregisterServer() {
        bool success;

        foreach (coclass; T) {
            static if (is(typeof(coclass.unregister))) {
                static assert(is(typeof(coclass.register)), "'unregister' must be matched by a corresponding 'register' in '" ~ coclass.stringof ~ "'.");

                success = unregisterComClass!(coclass)();
                coclass.unregister();
            }
        }

        return success ? S_OK : SELFREG_E_CLASS;
    }

}
