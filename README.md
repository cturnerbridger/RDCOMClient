# RDCOMClient (Updated Fork for R 4.6.x on Windows)

This repository is a **maintained public fork** of [`omegahat/RDCOMClient`](https://github.com/omegahat/RDCOMClient) (originally authored by Duncan Temple Lang), updated by **Christian** ([`@cturnerbridger`](https://github.com/cturnerbridger)) so that `RDCOMClient` compiles cleanly and runs reliably on **modern R for Windows (R 4.6.x)** using **Rtools 4.5 (GCC 14.3.0 UCRT)**.

> **Upstream attribution:** Original package and design by Duncan Temple Lang (`http://www.omegahat.net/RDCOMClient`, `https://github.com/omegahat/RDCOMClient`), licensed under GPL-2.

---

## Summary of Fixes in This Fork (`v0.96-2`)

### 1. R 4.6.x / Rtools 4.5 (GCC 14) Build & C++ Compatibility
- **Cleaned `src/Makevars.win`:** Removed obsolete hardcoded author paths (`d:/duncan/...`) and non-existent `.def` file references; streamlined `PKG_CPPFLAGS`, `PKG_CXXFLAGS`, and `PKG_LIBS` for modern Rtools toolchains.
- **Fixed C++ Header Linkage:** Moved R C headers (`<Rinternals.h>`, `<Rdefines.h>`, `RUtils.h`) outside `extern "C"` blocks across `src/connect.cpp`, `src/converters.cpp`, `src/RCOMObject.h`, and `src/RCOMObject.cpp` to resolve C++ template conflicts under GCC 14.
- **Updated R C API Macros:** Replaced legacy `EXTPTR_TAG` (removed in modern R internals) with `R_ExternalPtrTag`.

### 2. Full UTF-8 & Character Encoding Support
- **Dynamic UTF-8 / UTF-16 `BSTR` Conversion (`src/converters.cpp`):** Rewrote `AsBstr()` and `FromBstr()` to use `CP_UTF8` and dynamically allocate exact buffer lengths via `MultiByteToWideChar` and `WideCharToMultiByte`. Kept `wcslen()` in `FromBstr()` so `CoTaskMemAlloc`'d `LPOLESTR` strings (from `StringFromCLSID`) convert safely without crashing.
- **Proper R String Encoding (`src/RUtils.c`, `src/connect.cpp`, `src/converters.cpp`):** Marked returned strings with `Rf_mkCharCE(..., CE_UTF8)` and translated incoming R strings via `Rf_translateCharUTF8()` so non-UTF-8 strings (e.g., `latin1`) round-trip accurately through COM without corruption.
- **COM Server Member Lookup (`src/RCOMObject.cpp`):** Switched `WideCharToMultiByte` calls from `CP_ACP` to `CP_UTF8`.

### 3. Memory Leaks, Heap Corruption & Finalizer Safety
- **COM Reference-Counting & Resource Leaks:**
  - Released `IUnknown*` after `QueryInterface()` in `R_connect()`.
  - Released `FUNCDESC*` (`ReleaseFuncDesc`) and `ITypeInfo*` in `R_isReadOnly()`.
  - Freed `BSTR` arrays when `GetIDsOfNames()` fails in `R_COM_Invoke()`.
  - Removed duplicate `errorInfo->AddRef()` and freed `BSTR` source/description strings in `checkErrorInfo()` (`src/COMError.cpp`).
- **Heap Safety & Object Lifetime:**
  - Replaced `SysFreeString()` with `CoTaskMemFree()` for `StringFromCLSID()` buffers in `RCOMObject::QueryInterface()`.
  - Made `~RCOMObject()` virtual and re-enabled `delete this;` in `RCOMObject::Release()`.
  - Added `VariantInit()` before `VariantCopyInd()` in `createRVariantObject()`.
- **GC-Safe Finalizers & Package Unload:**
  - Updated `RDCOM_finalizer()` to inspect `R_ExternalPtrAddr()` directly instead of calling error-throwing helpers during R garbage collection.
  - Added `.onUnload()` in `R/zzz.R` to uninitialize COM and unload `RDCOMClient.dll` cleanly.

### 4. `SAFEARRAY` & Logical Array (`VARIANT_BOOL`) Fixes
- **Typed & Variant `SAFEARRAY` Extraction (`getArray` in `src/converters.cpp`):** Queried `SafeArrayGetVartype()` before calling `SafeArrayGetElement()` so both typed `SAFEARRAY`s (`VT_BOOL`, `VT_I4`, `VT_R8`, `VT_BSTR`, etc.) and `VT_VARIANT` arrays extract cleanly without corrupting the `VARIANT` tag, and added `VariantClear()` after converting each element to avoid leaking `BSTR`/`IDispatch*` copies.
- **1D & 2D Logical Arrays:** Updated `createRDCOMArray()` and `R_create2DArray()` to store proper 2-byte `VARIANT_BOOL` values (`VARIANT_TRUE` / `VARIANT_FALSE`) instead of C++ `bool` or 4-byte `int`.
- **2D String Matrices (`asCOMArray`):** Freed temporary `BSTR` allocations inside `R_create2DArray()`, removed redundant `SafeArrayAccessData()` locks around `SafeArrayPutElement()`, and registered `R_Variant_finalizer` instead of `RDCOM_SafeArray_finalizer` on the returned `VARIANT*` external pointer.

---

## Install

### Option 1: Install Prebuilt Windows Binary (Recommended — No Rtools Required)

Download or install `RDCOMClient_0.96-2.zip` directly from the [GitHub Releases](https://github.com/cturnerbridger/RDCOMClient/releases) page:

```r
install.packages(
  "https://github.com/cturnerbridger/RDCOMClient/releases/download/v0.96-2/RDCOMClient_0.96-2.zip",
  repos = NULL,
  type = "win.binary"
)
```

*(Built cleanly against R 4.6.x on Windows x64 UCRT.)*

### Option 2: Install from Source on GitHub

With **Rtools 4.5** installed (for R 4.6.x):

```r
# install.packages("remotes")
remotes::install_github("cturnerbridger/RDCOMClient")
```

Or from a local clone on the command line:

```sh
R CMD INSTALL --build RDCOMClient
```

---

## Usage Overview

This R package allows users to invoke (D)COM methods and access properties in any (D)COM object that implements the `IDispatch` interface. This includes Windows applications such as MS Excel, Word, PowerPoint, Outlook, and Windows Script Host components (`Scripting.Dictionary`, `Scripting.FileSystemObject`, etc.).

This is a **Windows-specific** package.

## Documentation

See the [`NEWS`](NEWS) file for full release notes, and check the original [Omegahat documentation](http://www.omegahat.net/RDCOMClient/) to learn more about package concepts and examples.

## Related Packages

+ `RDCOMServer`
+ `RDCOMEvents`
+ `SWinTypeLibs`
+ `SWinRegistry`
+ `rdcom`
+ `RExcel`
