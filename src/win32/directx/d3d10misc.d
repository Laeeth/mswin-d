/***********************************************************************\
*                               d3d10misc.d                             *
*                                                                       *
*                       Windows API header module                       *
*                                                                       *
*                       Placed into public domain                       *
\***********************************************************************/
module win32.directx.d3d10misc;

private import win32.windows;

private import win32.directx.dxgi;
private import win32.directx.d3d10;


enum D3D10_DRIVER_TYPE {
	D3D10_DRIVER_TYPE_HARDWARE  = 0,
	D3D10_DRIVER_TYPE_REFERENCE = 1,
	D3D10_DRIVER_TYPE_NULL      = 2,
	D3D10_DRIVER_TYPE_SOFTWARE  = 3
}

extern (C) const GUID GUID_DeviceType = {
	0xD722FB4D, 0x7A68, 0x437A, [0xB2, 0x0C, 0x58, 0x04, 0xEE, 0x24, 0x94, 0xA6]
};

extern(Windows) {
	HRESULT D3D10CreateDevice(IDXGIAdapter pAdapter,
	  D3D10_DRIVER_TYPE DriverType, HMODULE Software, UINT Flags,
	  UINT SDKVersion, ID3D10Device* ppDevice);
	HRESULT D3D10CreateDeviceAndSwapChain(IDXGIAdapter pAdapter,
	  D3D10_DRIVER_TYPE DriverType, HMODULE Software, UINT Flags,
	  UINT SDKVersion, DXGI_SWAP_CHAIN_DESC* pSwapChainDesc,
	  IDXGISwapChain* ppSwapChain, ID3D10Device* ppDevice);
	HRESULT D3D10CreateBlob(SIZE_T NumBytes, ID3D10Blob* ppBuffer);
}
