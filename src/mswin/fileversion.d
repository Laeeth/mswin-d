module mswin.fileversion;

import std.utf;

import win32.winver;

pragma(lib, "Version.lib");

/**
* Gets the version of the specified file, common for executables.
*/
Version fileVersion(string fileName)
{
    uint verSize = GetFileVersionInfoSizeW(std.utf.toUTF16z(fileName), null);
    if(!verSize) return Version();
    
    byte[] verData = new byte[verSize];
    uint valLen = 0;
    VS_FIXEDFILEINFO *fileInfo;
 
    if( !GetFileVersionInfoW(std.utf.toUTF16z(fileName), 0, verSize, cast(void*)verData.ptr) ||
        !VerQueryValueW( cast(void*)verData.ptr, "\\", cast(void**)&fileInfo, &valLen ) ||
        fileInfo.dwSignature != VS_FFI_SIGNATURE )
        return Version();

    return Version(
        (fileInfo.dwFileVersionMS & 0xFFFF0000) >> 16, fileInfo.dwFileVersionMS & 0x0000FFFF,
        (fileInfo.dwFileVersionLS & 0xFFFF0000) >> 16, fileInfo.dwFileVersionLS & 0x0000FFFF );
}


/**
* Represents a version number.
*/
struct Version {

    private int major_;
    private int minor_;
    private int build_ = -1;
    private int revision_ = -1;

    @property bool isNull() const {
        return major_ == 0 && minor_ == 0 &&
            build_ == -1 && revision_ == -1;
    }
 
    /**
    * Gets the value of the _major component.
    */
    @property int major() {
        return major_;
    }

    /**
    * Gets the value of the _minor component.
    */
    @property int minor() {
        return minor_;
    }

    /**
    * Gets the value of the _build component, -1 if not set
    */
    @property int build() {
        return build_;
    }

    /**
    * Gets the value of the _revision component, -1 if not set
    */
    @property int revision() {
        return revision_;
    }
    
    int opCmp(ref const Version v) {
        if (major_ != v.major_) {
            if (major_ > v.major_)
                return 1;
            return -1;
        }
        if (minor_ != v.minor_) {
            if (minor_ > v.minor_)
                return 1;
            return -1;
        }
        if (build_ != v.build_) {
            if (build_ > v.build_)
                return 1;
            return -1;
        }
        if (revision_ != v.revision_) {
            if (revision_ > v.revision_)
                return 1;
            return -1;
        }
        return 0;
    }

    bool opEquals(ref const Version v) {
        return (major_ == v.major_
                && minor_ == v.minor_
                && build_ == v.build_
                && revision_ == v.revision_);
    }

    hash_t toHash() {
        hash_t hash = (major_ & 0x0000000F) << 28;
        hash |= (minor_ & 0x000000FF) << 20;
        hash |= (build_ & 0x000000FF) << 12;
        hash |= revision_ & 0x00000FFF;
        return hash;
    }

     string toString() {
        import std.string;
        if(isNull)
            return "";
        string s = std.string.format("%d.%d", major_, minor_);
        if (build_ != -1) {
            s ~= std.string.format(".%d", build_);
            if (revision_ != -1)
                s ~= std.string.format(".%d", revision_);
        }
        return s;
    }

}