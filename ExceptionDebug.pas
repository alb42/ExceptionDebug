{**************************************************************************************************}
{                                                                                                  }
{ Project JEDI Code Library (JCL)                                                                  }
{                                                                                                  }
{ The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); }
{ you may not use this file except in compliance with the License. You may obtain a copy of the    }
{ License at http://www.mozilla.org/MPL/                                                           }
{                                                                                                  }
{ Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF   }
{ ANY KIND, either express or implied. See the License for the specific language governing rights  }
{ and limitations under the License.                                                               }
{                                                                                                  }
{ The Original Code is JclDebug.pas, JclPeImage.pas, JclTD32.pas, JclWin32.pas.                    }
{                                                                                                  }
{ The Initial Developers of the Original Code are Petr Vones and Marcel van Brakel.                }
{ Portions created by these individuals are Copyright (C) of these individuals.                    }
{ All Rights Reserved.                                                                             }
{                                                                                                  }
{ Contributor(s):                                                                                  }
{   Marcel van Brakel                                                                              }
{   Flier Lu (flier)                                                                               }
{   Florent Ouchet (outchy)                                                                        }
{   Robert Marquardt (marquardt)                                                                   }
{   Robert Rossmair (rrossmair)                                                                    }
{   Andreas Hausladen (ahuser)                                                                     }
{   Petr Vones (pvones)                                                                            }
{   Soeren Muehlbauer                                                                              }
{   Uwe Schuster (uschuster)                                                                       }
{   Matthias Thoma (mthoma)                                                                        }
{   Hallvard Vassbotn                                                                              }
{   Jean-Fabien Connault (cycocrew)                                                                }
{   Olivier Sannier (obones)                                                                       }
{   Heinz Zastrau (heinzz)                                                                         }
{   Andreas Hausladen (ahuser)                                                                     }
{   Peter Friese                                                                                   }
{                                                                                                  }
{**************************************************************************************************}

unit ExceptionDebug;

interface

uses
  Winapi.Windows, Winapi.AccCtrl, Winapi.ActiveX, TypInfo, Math, Winapi.TLHelp32, Winapi.PsApi,
  System.Classes, System.SysUtils, System.Contnrs, SyncObjs, AnsiStrings;


type
  PLOADED_IMAGE = ^LOADED_IMAGE;
  {$EXTERNALSYM PLOADED_IMAGE}
  _LOADED_IMAGE = record
    ModuleName: PAnsiChar;
    hFile: THandle;
    MappedAddress: PUCHAR;
    FileHeader: PImageNtHeaders;
    LastRvaSection: PImageSectionHeader;
    NumberOfSections: ULONG;
    Sections: PImageSectionHeader;
    Characteristics: ULONG;
    fSystemImage: ByteBool;
    fDOSImage: ByteBool;
    Links: LIST_ENTRY;
    SizeOfImage: ULONG;
  end;
  {$EXTERNALSYM _LOADED_IMAGE}
  LOADED_IMAGE = _LOADED_IMAGE;
  {$EXTERNALSYM LOADED_IMAGE}
  TLoadedImage = LOADED_IMAGE;
  PLoadedImage = PLOADED_IMAGE;

  PIMAGE_SYMBOL = ^IMAGE_SYMBOL;
  {$EXTERNALSYM PIMAGE_SYMBOL}
  _IMAGE_SYMBOL = packed record  // MUST pack to obtain the right size
    Name: array [0..7] of AnsiChar;
    Value: ULONG;
    SectionNumber: USHORT;
    _Type: USHORT;
    StorageClass: BYTE;
    NumberOfAuxSymbols: BYTE;
  end;
  {$EXTERNALSYM _IMAGE_SYMBOL}
  IMAGE_SYMBOL = _IMAGE_SYMBOL;
  {$EXTERNALSYM IMAGE_SYMBOL}
  TImageSymbol = IMAGE_SYMBOL;
  PImageSymbol = PIMAGE_SYMBOL;

//
// Define checksum function prototypes.
//

function CheckSumMappedFile(BaseAddress: Pointer; FileLength: DWORD;
  out HeaderSum, CheckSum: DWORD): PImageNtHeaders; stdcall;
{$EXTERNALSYM CheckSumMappedFile}

// line 227

function GetImageUnusedHeaderBytes(const LoadedImage: LOADED_IMAGE;
  var SizeUnusedHeaderBytes: DWORD): DWORD; stdcall;
{$EXTERNALSYM GetImageUnusedHeaderBytes}

// line 285

function MapAndLoad(ImageName, DllPath: PAnsiChar; var LoadedImage: LOADED_IMAGE;
  DotDll: BOOL; ReadOnly: BOOL): BOOL; stdcall;
{$EXTERNALSYM MapAndLoad}

function UnMapAndLoad(const LoadedImage: LOADED_IMAGE): BOOL; stdcall;
{$EXTERNALSYM UnMapAndLoad}

function TouchFileTimes(const FileHandle: THandle; const pSystemTime: TSystemTime): BOOL; stdcall;
{$EXTERNALSYM TouchFileTimes}

// line 347

function ImageDirectoryEntryToData(Base: Pointer; MappedAsImage: ByteBool;
  DirectoryEntry: USHORT; var Size: ULONG): Pointer; stdcall;
{$EXTERNALSYM ImageDirectoryEntryToData}

function ImageRvaToSection(NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG): PImageSectionHeader; stdcall;
{$EXTERNALSYM ImageRvaToSection}

function ImageRvaToVa(NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG;
  LastRvaSection: PPImageSectionHeader): Pointer; stdcall;
{$EXTERNALSYM ImageRvaToVa}


// line 461

//
// UnDecorateSymbolName Flags
//

const
  UNDNAME_COMPLETE               = ($0000); // Enable full undecoration
  {$EXTERNALSYM UNDNAME_COMPLETE}
  UNDNAME_NO_LEADING_UNDERSCORES = ($0001); // Remove leading underscores from MS extended keywords
  {$EXTERNALSYM UNDNAME_NO_LEADING_UNDERSCORES}
  UNDNAME_NO_MS_KEYWORDS         = ($0002); // Disable expansion of MS extended keywords
  {$EXTERNALSYM UNDNAME_NO_MS_KEYWORDS}
  UNDNAME_NO_FUNCTION_RETURNS    = ($0004); // Disable expansion of return type for primary declaration
  {$EXTERNALSYM UNDNAME_NO_FUNCTION_RETURNS}
  UNDNAME_NO_ALLOCATION_MODEL    = ($0008); // Disable expansion of the declaration model
  {$EXTERNALSYM UNDNAME_NO_ALLOCATION_MODEL}
  UNDNAME_NO_ALLOCATION_LANGUAGE = ($0010); // Disable expansion of the declaration language specifier
  {$EXTERNALSYM UNDNAME_NO_ALLOCATION_LANGUAGE}
  UNDNAME_NO_MS_THISTYPE         = ($0020); // NYI Disable expansion of MS keywords on the 'this' type for primary declaration
  {$EXTERNALSYM UNDNAME_NO_MS_THISTYPE}
  UNDNAME_NO_CV_THISTYPE         = ($0040); // NYI Disable expansion of CV modifiers on the 'this' type for primary declaration
  {$EXTERNALSYM UNDNAME_NO_CV_THISTYPE}
  UNDNAME_NO_THISTYPE            = ($0060); // Disable all modifiers on the 'this' type
  {$EXTERNALSYM UNDNAME_NO_THISTYPE}
  UNDNAME_NO_ACCESS_SPECIFIERS   = ($0080); // Disable expansion of access specifiers for members
  {$EXTERNALSYM UNDNAME_NO_ACCESS_SPECIFIERS}
  UNDNAME_NO_THROW_SIGNATURES    = ($0100); // Disable expansion of 'throw-signatures' for functions and pointers to functions
  {$EXTERNALSYM UNDNAME_NO_THROW_SIGNATURES}
  UNDNAME_NO_MEMBER_TYPE         = ($0200); // Disable expansion of 'static' or 'virtual'ness of members
  {$EXTERNALSYM UNDNAME_NO_MEMBER_TYPE}
  UNDNAME_NO_RETURN_UDT_MODEL    = ($0400); // Disable expansion of MS model for UDT returns
  {$EXTERNALSYM UNDNAME_NO_RETURN_UDT_MODEL}
  UNDNAME_32_BIT_DECODE          = ($0800); // Undecorate 32-bit decorated names
  {$EXTERNALSYM UNDNAME_32_BIT_DECODE}
  UNDNAME_NAME_ONLY              = ($1000); // Crack only the name for primary declaration;
  {$EXTERNALSYM UNDNAME_NAME_ONLY}
                                                                                                   //  return just [scope::]name.  Does expand template params
  UNDNAME_NO_ARGUMENTS    = ($2000); // Don't undecorate arguments to function
  {$EXTERNALSYM UNDNAME_NO_ARGUMENTS}
  UNDNAME_NO_SPECIAL_SYMS = ($4000); // Don't undecorate special names (v-table, vcall, vector xxx, metatype, etc)
  {$EXTERNALSYM UNDNAME_NO_SPECIAL_SYMS}

// line 1342

type
  {$EXTERNALSYM SYM_TYPE}
  SYM_TYPE = (
    SymNone,
    SymCoff,
    SymCv,
    SymPdb,
    SymExport,
    SymDeferred,
    SymSym                  { .sym file }
  );
  TSymType = SYM_TYPE;

  { symbol data structure }
  {$EXTERNALSYM PImagehlpSymbolA}
  PImagehlpSymbolA = ^TImagehlpSymbolA;
  {$EXTERNALSYM _IMAGEHLP_SYMBOLA}
  _IMAGEHLP_SYMBOLA = packed record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_SYMBOL) }
    Address: DWORD;                                     { virtual address including dll base address }
    Size: DWORD;                                        { estimated size of symbol, can be zero }
    Flags: DWORD;                                       { info about the symbols, see the SYMF defines }
    MaxNameLength: DWORD;                               { maximum size of symbol name in 'Name' }
    Name: packed array[0..0] of AnsiChar;               { symbol name (null terminated string) }
  end;
  {$EXTERNALSYM IMAGEHLP_SYMBOLA}
  IMAGEHLP_SYMBOLA = _IMAGEHLP_SYMBOLA;
  {$EXTERNALSYM TImagehlpSymbolA}
  TImagehlpSymbolA = _IMAGEHLP_SYMBOLA;

  { symbol data structure }
  {$EXTERNALSYM PImagehlpSymbolA64}
  PImagehlpSymbolA64 = ^TImagehlpSymbolA64;
  {$EXTERNALSYM _IMAGEHLP_SYMBOLA64}
  _IMAGEHLP_SYMBOLA64 = packed record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_SYMBOL) }
    Address: Int64;                                { virtual address including dll base address }
    Size: DWORD;                                        { estimated size of symbol, can be zero }
    Flags: DWORD;                                       { info about the symbols, see the SYMF defines }
    MaxNameLength: DWORD;                               { maximum size of symbol name in 'Name' }
    Name: packed array[0..0] of AnsiChar;               { symbol name (null terminated string) }
  end;
  {$EXTERNALSYM IMAGEHLP_SYMBOLA64}
  IMAGEHLP_SYMBOLA64 = _IMAGEHLP_SYMBOLA64;
  {$EXTERNALSYM TImagehlpSymbolA64}
  TImagehlpSymbolA64 = _IMAGEHLP_SYMBOLA64;

  { symbol data structure }
  {$EXTERNALSYM PImagehlpSymbolW}
  PImagehlpSymbolW = ^TImagehlpSymbolW;
  {$EXTERNALSYM _IMAGEHLP_SYMBOLW}
  _IMAGEHLP_SYMBOLW = packed record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_SYMBOL) }
    Address: DWORD;                                     { virtual address including dll base address }
    Size: DWORD;                                        { estimated size of symbol, can be zero }
    Flags: DWORD;                                       { info about the symbols, see the SYMF defines }
    MaxNameLength: DWORD;                               { maximum size of symbol name in 'Name' }
    Name: packed array[0..0] of WideChar;               { symbol name (null terminated string) }
  end;
  {$EXTERNALSYM IMAGEHLP_SYMBOLW}
  IMAGEHLP_SYMBOLW = _IMAGEHLP_SYMBOLW;
  {$EXTERNALSYM TImagehlpSymbolW}
  TImagehlpSymbolW = _IMAGEHLP_SYMBOLW;

  { symbol data structure }
  {$EXTERNALSYM PImagehlpSymbolW64}
  PImagehlpSymbolW64 = ^TImagehlpSymbolW64;
  {$EXTERNALSYM _IMAGEHLP_SYMBOLW64}
  _IMAGEHLP_SYMBOLW64 = packed record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_SYMBOL) }
    Address: Int64;                                { virtual address including dll base address }
    Size: DWORD;                                        { estimated size of symbol, can be zero }
    Flags: DWORD;                                       { info about the symbols, see the SYMF defines }
    MaxNameLength: DWORD;                               { maximum size of symbol name in 'Name' }
    Name: packed array[0..0] of WideChar;               { symbol name (null terminated string) }
  end;
  {$EXTERNALSYM IMAGEHLP_SYMBOLW64}
  IMAGEHLP_SYMBOLW64 = _IMAGEHLP_SYMBOLW64;
  {$EXTERNALSYM TImagehlpSymbolW64}
  TImagehlpSymbolW64 = _IMAGEHLP_SYMBOLW64;

  { module data structure }
  {$EXTERNALSYM PImagehlpModuleA}
  PImagehlpModuleA = ^TImagehlpModuleA;
  {$EXTERNALSYM _IMAGEHLP_MODULEA}
  _IMAGEHLP_MODULEA = record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_MODULE) }
    BaseOfImage: DWORD;                                 { base load address of module }
    ImageSize: DWORD;                                   { virtual size of the loaded module }
    TimeDateStamp: DWORD;                               { date/time stamp from pe header }
    CheckSum: DWORD;                                    { checksum from the pe header }
    NumSyms: DWORD;                                     { number of symbols in the symbol table }
    SymType: TSymType;                                  { type of symbols loaded }
    ModuleName: packed array[0..31] of AnsiChar;        { module name }
    ImageName: packed array[0..255] of AnsiChar;        { image name }
    LoadedImageName: packed array[0..255] of AnsiChar;  { symbol file name }
  end;
  {$EXTERNALSYM IMAGEHLP_MODULEA}
  IMAGEHLP_MODULEA = _IMAGEHLP_MODULEA;
  {$EXTERNALSYM TImagehlpModuleA}
  TImagehlpModuleA = _IMAGEHLP_MODULEA;

  { module data structure }
  {$EXTERNALSYM PImagehlpModuleA64}
  PImagehlpModuleA64 = ^TImagehlpModuleA64;
  {$EXTERNALSYM _IMAGEHLP_MODULEA64}
  _IMAGEHLP_MODULEA64 = record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_MODULE) }
    BaseOfImage: Int64;                            { base load address of module }
    ImageSize: DWORD;                                   { virtual size of the loaded module }
    TimeDateStamp: DWORD;                               { date/time stamp from pe header }
    CheckSum: DWORD;                                    { checksum from the pe header }
    NumSyms: DWORD;                                     { number of symbols in the symbol table }
    SymType: TSymType;                                  { type of symbols loaded }
    ModuleName: packed array[0..31] of AnsiChar;        { module name }
    ImageName: packed array[0..255] of AnsiChar;        { image name }
    LoadedImageName: packed array[0..255] of AnsiChar;  { symbol file name }
  end;
  {$EXTERNALSYM IMAGEHLP_MODULEA64}
  IMAGEHLP_MODULEA64 = _IMAGEHLP_MODULEA64;
  {$EXTERNALSYM TImagehlpModuleA64}
  TImagehlpModuleA64 = _IMAGEHLP_MODULEA64;

  { module data structure }
  {$EXTERNALSYM PImagehlpModuleW}
  PImagehlpModuleW = ^TImagehlpModuleW;
  {$EXTERNALSYM _IMAGEHLP_MODULEW}
  _IMAGEHLP_MODULEW = record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_MODULE) }
    BaseOfImage: DWORD;                                 { base load address of module }
    ImageSize: DWORD;                                   { virtual size of the loaded module }
    TimeDateStamp: DWORD;                               { date/time stamp from pe header }
    CheckSum: DWORD;                                    { checksum from the pe header }
    NumSyms: DWORD;                                     { number of symbols in the symbol table }
    SymType: TSymType;                                  { type of symbols loaded }
    ModuleName: packed array[0..31] of WideChar;        { module name }
    ImageName: packed array[0..255] of WideChar;        { image name }
    LoadedImageName: packed array[0..255] of WideChar;  { symbol file name }
  end;
  {$EXTERNALSYM IMAGEHLP_MODULEW}
  IMAGEHLP_MODULEW = _IMAGEHLP_MODULEW;
  {$EXTERNALSYM TImagehlpModuleW}
  TImagehlpModuleW = _IMAGEHLP_MODULEW;

  { module data structure }
  {$EXTERNALSYM PImagehlpModuleW64}
  PImagehlpModuleW64 = ^TImagehlpModuleW64;
  {$EXTERNALSYM _IMAGEHLP_MODULEW64}
  _IMAGEHLP_MODULEW64 = record
    SizeOfStruct: DWORD;                                { set to sizeof(IMAGEHLP_MODULE) }
    BaseOfImage: Int64;                            { base load address of module }
    ImageSize: DWORD;                                   { virtual size of the loaded module }
    TimeDateStamp: DWORD;                               { date/time stamp from pe header }
    CheckSum: DWORD;                                    { checksum from the pe header }
    NumSyms: DWORD;                                     { number of symbols in the symbol table }
    SymType: TSymType;                                  { type of symbols loaded }
    ModuleName: packed array[0..31] of WideChar;        { module name }
    ImageName: packed array[0..255] of WideChar;        { image name }
    LoadedImageName: packed array[0..255] of WideChar;  { symbol file name }
  end;
  {$EXTERNALSYM IMAGEHLP_MODULEW64}
  IMAGEHLP_MODULEW64 = _IMAGEHLP_MODULEW64;
  {$EXTERNALSYM TImagehlpModuleW64}
  TImagehlpModuleW64 = _IMAGEHLP_MODULEW64;

  _IMAGEHLP_LINEA = packed record
    SizeOfStruct: DWORD;           // set to sizeof(IMAGEHLP_LINE)
    Key: Pointer;                  // internal
    LineNumber: DWORD;             // line number in file
    FileName: PAnsiChar;           // full filename
    Address: DWORD;                // first instruction of line
  end;
  IMAGEHLP_LINEA = _IMAGEHLP_LINEA;
  PIMAGEHLP_LINEA = ^_IMAGEHLP_LINEA;
  TImageHlpLineA = _IMAGEHLP_LINEA;
  PImageHlpLineA = PIMAGEHLP_LINEA;

  _IMAGEHLP_LINEA64 = packed record
    SizeOfStruct: DWORD;           // set to sizeof(IMAGEHLP_LINE)
    Key: Pointer;                  // internal
    LineNumber: DWORD;             // line number in file
    FileName: PAnsiChar;           // full filename
    Address: Int64;           // first instruction of line
  end;
  IMAGEHLP_LINEA64 = _IMAGEHLP_LINEA64;
  PIMAGEHLP_LINEA64 = ^_IMAGEHLP_LINEA64;
  TImageHlpLineA64 = _IMAGEHLP_LINEA64;
  PImageHlpLineA64 = PIMAGEHLP_LINEA64;

  _IMAGEHLP_LINEW = packed record
    SizeOfStruct: DWORD;           // set to sizeof(IMAGEHLP_LINE)
    Key: Pointer;                  // internal
    LineNumber: DWORD;             // line number in file
    FileName: PWideChar;           // full filename
    Address: DWORD;                // first instruction of line
  end;
  IMAGEHLP_LINEW = _IMAGEHLP_LINEW;
  PIMAGEHLP_LINEW = ^_IMAGEHLP_LINEW;
  TImageHlpLineW = _IMAGEHLP_LINEW;
  PImageHlpLineW = PIMAGEHLP_LINEW;

  _IMAGEHLP_LINEW64 = packed record
    SizeOfStruct: DWORD;           // set to sizeof(IMAGEHLP_LINE)
    Key: Pointer;                  // internal
    LineNumber: DWORD;             // line number in file
    FileName: PWideChar;           // full filename
    Address: Int64;           // first instruction of line
  end;
  IMAGEHLP_LINEW64 = _IMAGEHLP_LINEW64;
  PIMAGEHLP_LINEW64 = ^_IMAGEHLP_LINEW64;
  TImageHlpLineW64 = _IMAGEHLP_LINEW64;
  PImageHlpLineW64 = PIMAGEHLP_LINEW64;

// line 1475

//
// options that are set/returned by SymSetOptions() & SymGetOptions()
// these are used as a mask
//

const
  SYMOPT_CASE_INSENSITIVE       = $00000001;
  {$EXTERNALSYM SYMOPT_CASE_INSENSITIVE}
  SYMOPT_UNDNAME                = $00000002;
  {$EXTERNALSYM SYMOPT_UNDNAME}
  SYMOPT_DEFERRED_LOADS         = $00000004;
  {$EXTERNALSYM SYMOPT_DEFERRED_LOADS}
  SYMOPT_NO_CPP                 = $00000008;
  {$EXTERNALSYM SYMOPT_NO_CPP}
  SYMOPT_LOAD_LINES             = $00000010;
  {$EXTERNALSYM SYMOPT_LOAD_LINES}
  SYMOPT_OMAP_FIND_NEAREST      = $00000020;
  {$EXTERNALSYM SYMOPT_OMAP_FIND_NEAREST}
  SYMOPT_LOAD_ANYTHING          = $00000040;
  {$EXTERNALSYM SYMOPT_LOAD_ANYTHING}
  SYMOPT_IGNORE_CVREC           = $00000080;
  {$EXTERNALSYM SYMOPT_IGNORE_CVREC}
  SYMOPT_NO_UNQUALIFIED_LOADS   = $00000100;
  {$EXTERNALSYM SYMOPT_NO_UNQUALIFIED_LOADS}
  SYMOPT_FAIL_CRITICAL_ERRORS   = $00000200;
  {$EXTERNALSYM SYMOPT_FAIL_CRITICAL_ERRORS}
  SYMOPT_EXACT_SYMBOLS          = $00000400;
  {$EXTERNALSYM SYMOPT_EXACT_SYMBOLS}
  SYMOPT_ALLOW_ABSOLUTE_SYMBOLS = $00000800;
  {$EXTERNALSYM SYMOPT_ALLOW_ABSOLUTE_SYMBOLS}
  SYMOPT_IGNORE_NT_SYMPATH      = $00001000;
  {$EXTERNALSYM SYMOPT_IGNORE_NT_SYMPATH}
  SYMOPT_INCLUDE_32BIT_MODULES  = $00002000;
  {$EXTERNALSYM SYMOPT_INCLUDE_32BIT_MODULES}
  SYMOPT_PUBLICS_ONLY           = $00004000;
  {$EXTERNALSYM SYMOPT_PUBLICS_ONLY}
  SYMOPT_NO_PUBLICS             = $00008000;
  {$EXTERNALSYM SYMOPT_NO_PUBLICS}
  SYMOPT_AUTO_PUBLICS           = $00010000;
  {$EXTERNALSYM SYMOPT_AUTO_PUBLICS}
  SYMOPT_NO_IMAGE_SEARCH        = $00020000;
  {$EXTERNALSYM SYMOPT_NO_IMAGE_SEARCH}
  SYMOPT_SECURE                 = $00040000;
  {$EXTERNALSYM SYMOPT_SECURE}
  SYMOPT_NO_PROMPTS             = $00080000;
  {$EXTERNALSYM SYMOPT_NO_PROMPTS}

  SYMOPT_DEBUG                  = $80000000;
  {$EXTERNALSYM SYMOPT_DEBUG}

// IoAPI.h


const
  NERR_Success = 0; // Success
  {$EXTERNALSYM NERR_Success}

// ERROR_ equates can be intermixed with NERR_ equates.

//    NERR_BASE is the base of error codes from network utilities,
//      chosen to avoid conflict with system and redirector error codes.
//      2100 is a value that has been assigned to us by system.

  NERR_BASE = 2100;
  {$EXTERNALSYM NERR_BASE}


//*INTERNAL_ONLY*

{**********WARNING *****************
 *See the comment in lmcons.h for  *
 *info on the allocation of errors *
 ***********************************}

{**********WARNING *****************
 *The range 2750-2799 has been     *
 *allocated to the IBM LAN Server  *
 ***********************************}

{**********WARNING *****************
 *The range 2900-2999 has been     *
 *reserved for Microsoft OEMs      *
 ***********************************}

// UNUSED BASE+0
// UNUSED BASE+1
  NERR_NetNotStarted = (NERR_BASE+2); // The workstation driver is not installed.
  {$EXTERNALSYM NERR_NetNotStarted}
  NERR_UnknownServer = (NERR_BASE+3); // The server could not be located.
  {$EXTERNALSYM NERR_UnknownServer}
  NERR_ShareMem      = (NERR_BASE+4); // An internal error occurred.  The network cannot access a shared memory segment.
  {$EXTERNALSYM NERR_ShareMem}

  NERR_NoNetworkResource = (NERR_BASE+5); // A network resource shortage occurred .
  {$EXTERNALSYM NERR_NoNetworkResource}
  NERR_RemoteOnly        = (NERR_BASE+6); // This operation is not supported on workstations.
  {$EXTERNALSYM NERR_RemoteOnly}
  NERR_DevNotRedirected  = (NERR_BASE+7); // The device is not connected.
  {$EXTERNALSYM NERR_DevNotRedirected}
// NERR_BASE+8 is used for ERROR_CONNECTED_OTHER_PASSWORD
// NERR_BASE+9 is used for ERROR_CONNECTED_OTHER_PASSWORD_DEFAULT
// UNUSED BASE+10
// UNUSED BASE+11
// UNUSED BASE+12
// UNUSED BASE+13
  NERR_ServerNotStarted = (NERR_BASE+14); // The Server service is not started.
  {$EXTERNALSYM NERR_ServerNotStarted}
  NERR_ItemNotFound     = (NERR_BASE+15); // The queue is empty.
  {$EXTERNALSYM NERR_ItemNotFound}
  NERR_UnknownDevDir    = (NERR_BASE+16); // The device or directory does not exist.
  {$EXTERNALSYM NERR_UnknownDevDir}
  NERR_RedirectedPath   = (NERR_BASE+17); // The operation is invalid on a redirected resource.
  {$EXTERNALSYM NERR_RedirectedPath}
  NERR_DuplicateShare   = (NERR_BASE+18); // The name has already been shared.
  {$EXTERNALSYM NERR_DuplicateShare}
  NERR_NoRoom           = (NERR_BASE+19); // The server is currently out of the requested resource.
  {$EXTERNALSYM NERR_NoRoom}
// UNUSED BASE+20
  NERR_TooManyItems    = (NERR_BASE+21); // Requested addition of items exceeds the maximum allowed.
  {$EXTERNALSYM NERR_TooManyItems}
  NERR_InvalidMaxUsers = (NERR_BASE+22); // The Peer service supports only two simultaneous users.
  {$EXTERNALSYM NERR_InvalidMaxUsers}
  NERR_BufTooSmall     = (NERR_BASE+23); // The API return buffer is too small.
  {$EXTERNALSYM NERR_BufTooSmall}
// UNUSED BASE+24
// UNUSED BASE+25
// UNUSED BASE+26
  NERR_RemoteErr = (NERR_BASE+27); // A remote API error occurred.
  {$EXTERNALSYM NERR_RemoteErr}
// UNUSED BASE+28
// UNUSED BASE+29
// UNUSED BASE+30
  NERR_LanmanIniError = (NERR_BASE+31); // An error occurred when opening or reading the configuration file.
  {$EXTERNALSYM NERR_LanmanIniError}
// UNUSED BASE+32
// UNUSED BASE+33
// UNUSED BASE+34
// UNUSED BASE+35
  NERR_NetworkError           = (NERR_BASE+36); // A general network error occurred.
  {$EXTERNALSYM NERR_NetworkError}
  NERR_WkstaInconsistentState = (NERR_BASE+37);
  {$EXTERNALSYM NERR_WkstaInconsistentState}
    // The Workstation service is in an inconsistent state. Restart the computer before restarting the Workstation service.
  NERR_WkstaNotStarted   = (NERR_BASE+38); // The Workstation service has not been started.
  {$EXTERNALSYM NERR_WkstaNotStarted}
  NERR_BrowserNotStarted = (NERR_BASE+39); // The requested information is not available.
  {$EXTERNALSYM NERR_BrowserNotStarted}
  NERR_InternalError     = (NERR_BASE+40); // An internal Windows 2000 error occurred.
  {$EXTERNALSYM NERR_InternalError}
  NERR_BadTransactConfig = (NERR_BASE+41); // The server is not configured for transactions.
  {$EXTERNALSYM NERR_BadTransactConfig}
  NERR_InvalidAPI        = (NERR_BASE+42); // The requested API is not supported on the remote server.
  {$EXTERNALSYM NERR_InvalidAPI}
  NERR_BadEventName      = (NERR_BASE+43); // The event name is invalid.
  {$EXTERNALSYM NERR_BadEventName}
  NERR_DupNameReboot     = (NERR_BASE+44); // The computer name already exists on the network. Change it and restart the computer.
  {$EXTERNALSYM NERR_DupNameReboot}

//
//      Config API related
//              Error codes from BASE+45 to BASE+49


// UNUSED BASE+45
  NERR_CfgCompNotFound  = (NERR_BASE+46); // The specified component could not be found in the configuration information.
  {$EXTERNALSYM NERR_CfgCompNotFound}
  NERR_CfgParamNotFound = (NERR_BASE+47); // The specified parameter could not be found in the configuration information.
  {$EXTERNALSYM NERR_CfgParamNotFound}
  NERR_LineTooLong = (NERR_BASE+49); // A line in the configuration file is too long.
  {$EXTERNALSYM NERR_LineTooLong}

//
//      Spooler API related
//              Error codes from BASE+50 to BASE+79


  NERR_QNotFound        = (NERR_BASE+50); // The printer does not exist.
  {$EXTERNALSYM NERR_QNotFound}
  NERR_JobNotFound      = (NERR_BASE+51); // The print job does not exist.
  {$EXTERNALSYM NERR_JobNotFound}
  NERR_DestNotFound     = (NERR_BASE+52); // The printer destination cannot be found.
  {$EXTERNALSYM NERR_DestNotFound}
  NERR_DestExists       = (NERR_BASE+53); // The printer destination already exists.
  {$EXTERNALSYM NERR_DestExists}
  NERR_QExists          = (NERR_BASE+54); // The printer queue already exists.
  {$EXTERNALSYM NERR_QExists}
  NERR_QNoRoom          = (NERR_BASE+55); // No more printers can be added.
  {$EXTERNALSYM NERR_QNoRoom}
  NERR_JobNoRoom        = (NERR_BASE+56); // No more print jobs can be added.
  {$EXTERNALSYM NERR_JobNoRoom}
  NERR_DestNoRoom       = (NERR_BASE+57); // No more printer destinations can be added.
  {$EXTERNALSYM NERR_DestNoRoom}
  NERR_DestIdle         = (NERR_BASE+58); // This printer destination is idle and cannot accept control operations.
  {$EXTERNALSYM NERR_DestIdle}
  NERR_DestInvalidOp    = (NERR_BASE+59); // This printer destination request contains an invalid control function.
  {$EXTERNALSYM NERR_DestInvalidOp}
  NERR_ProcNoRespond    = (NERR_BASE+60); // The print processor is not responding.
  {$EXTERNALSYM NERR_ProcNoRespond}
  NERR_SpoolerNotLoaded = (NERR_BASE+61); // The spooler is not running.
  {$EXTERNALSYM NERR_SpoolerNotLoaded}
  NERR_DestInvalidState = (NERR_BASE+62); // This operation cannot be performed on the print destination in its current state.
  {$EXTERNALSYM NERR_DestInvalidState}
  NERR_QInvalidState    = (NERR_BASE+63); // This operation cannot be performed on the printer queue in its current state.
  {$EXTERNALSYM NERR_QInvalidState}
  NERR_JobInvalidState  = (NERR_BASE+64); // This operation cannot be performed on the print job in its current state.
  {$EXTERNALSYM NERR_JobInvalidState}
  NERR_SpoolNoMemory    = (NERR_BASE+65); // A spooler memory allocation failure occurred.
  {$EXTERNALSYM NERR_SpoolNoMemory}
  NERR_DriverNotFound   = (NERR_BASE+66); // The device driver does not exist.
  {$EXTERNALSYM NERR_DriverNotFound}
  NERR_DataTypeInvalid  = (NERR_BASE+67); // The data type is not supported by the print processor.
  {$EXTERNALSYM NERR_DataTypeInvalid}
  NERR_ProcNotFound     = (NERR_BASE+68); // The print processor is not installed.
  {$EXTERNALSYM NERR_ProcNotFound}

//
//      Service API related
//              Error codes from BASE+80 to BASE+99


  NERR_ServiceTableLocked  = (NERR_BASE+80); // The service database is locked.
  {$EXTERNALSYM NERR_ServiceTableLocked}
  NERR_ServiceTableFull    = (NERR_BASE+81); // The service table is full.
  {$EXTERNALSYM NERR_ServiceTableFull}
  NERR_ServiceInstalled    = (NERR_BASE+82); // The requested service has already been started.
  {$EXTERNALSYM NERR_ServiceInstalled}
  NERR_ServiceEntryLocked  = (NERR_BASE+83); // The service does not respond to control actions.
  {$EXTERNALSYM NERR_ServiceEntryLocked}
  NERR_ServiceNotInstalled = (NERR_BASE+84); // The service has not been started.
  {$EXTERNALSYM NERR_ServiceNotInstalled}
  NERR_BadServiceName      = (NERR_BASE+85); // The service name is invalid.
  {$EXTERNALSYM NERR_BadServiceName}
  NERR_ServiceCtlTimeout   = (NERR_BASE+86); // The service is not responding to the control function.
  {$EXTERNALSYM NERR_ServiceCtlTimeout}
  NERR_ServiceCtlBusy      = (NERR_BASE+87); // The service control is busy.
  {$EXTERNALSYM NERR_ServiceCtlBusy}
  NERR_BadServiceProgName  = (NERR_BASE+88); // The configuration file contains an invalid service program name.
  {$EXTERNALSYM NERR_BadServiceProgName}
  NERR_ServiceNotCtrl      = (NERR_BASE+89); // The service could not be controlled in its present state.
  {$EXTERNALSYM NERR_ServiceNotCtrl}
  NERR_ServiceKillProc     = (NERR_BASE+90); // The service ended abnormally.
  {$EXTERNALSYM NERR_ServiceKillProc}
  NERR_ServiceCtlNotValid  = (NERR_BASE+91); // The requested pause,continue, or stop is not valid for this service.
  {$EXTERNALSYM NERR_ServiceCtlNotValid}
  NERR_NotInDispatchTbl    = (NERR_BASE+92); // The service control dispatcher could not find the service name in the dispatch table.
  {$EXTERNALSYM NERR_NotInDispatchTbl}
  NERR_BadControlRecv      = (NERR_BASE+93); // The service control dispatcher pipe read failed.
  {$EXTERNALSYM NERR_BadControlRecv}
  NERR_ServiceNotStarting  = (NERR_BASE+94); // A thread for the new service could not be created.
  {$EXTERNALSYM NERR_ServiceNotStarting}

//
//      Wksta and Logon API related
//              Error codes from BASE+100 to BASE+118


  NERR_AlreadyLoggedOn   = (NERR_BASE+100); // This workstation is already logged on to the local-area network.
  {$EXTERNALSYM NERR_AlreadyLoggedOn}
  NERR_NotLoggedOn       = (NERR_BASE+101); // The workstation is not logged on to the local-area network.
  {$EXTERNALSYM NERR_NotLoggedOn}
  NERR_BadUsername       = (NERR_BASE+102); // The user name or group name parameter is invalid.
  {$EXTERNALSYM NERR_BadUsername}
  NERR_BadPassword       = (NERR_BASE+103); // The password parameter is invalid.
  {$EXTERNALSYM NERR_BadPassword}
  NERR_UnableToAddName_W = (NERR_BASE+104); // @W The logon processor did not add the message alias.
  {$EXTERNALSYM NERR_UnableToAddName_W}
  NERR_UnableToAddName_F = (NERR_BASE+105); // The logon processor did not add the message alias.
  {$EXTERNALSYM NERR_UnableToAddName_F}
  NERR_UnableToDelName_W = (NERR_BASE+106); // @W The logoff processor did not delete the message alias.
  {$EXTERNALSYM NERR_UnableToDelName_W}
  NERR_UnableToDelName_F = (NERR_BASE+107); // The logoff processor did not delete the message alias.
  {$EXTERNALSYM NERR_UnableToDelName_F}
// UNUSED BASE+108
  NERR_LogonsPaused        = (NERR_BASE+109); // Network logons are paused.
  {$EXTERNALSYM NERR_LogonsPaused}
  NERR_LogonServerConflict = (NERR_BASE+110); // A centralized logon-server conflict occurred.
  {$EXTERNALSYM NERR_LogonServerConflict}
  NERR_LogonNoUserPath     = (NERR_BASE+111); // The server is configured without a valid user path.
  {$EXTERNALSYM NERR_LogonNoUserPath}
  NERR_LogonScriptError    = (NERR_BASE+112); // An error occurred while loading or running the logon script.
  {$EXTERNALSYM NERR_LogonScriptError}
// UNUSED BASE+113
  NERR_StandaloneLogon     = (NERR_BASE+114); // The logon server was not specified.  Your computer will be logged on as STANDALONE.
  {$EXTERNALSYM NERR_StandaloneLogon}
  NERR_LogonServerNotFound = (NERR_BASE+115); // The logon server could not be found.
  {$EXTERNALSYM NERR_LogonServerNotFound}
  NERR_LogonDomainExists   = (NERR_BASE+116); // There is already a logon domain for this computer.
  {$EXTERNALSYM NERR_LogonDomainExists}
  NERR_NonValidatedLogon   = (NERR_BASE+117); // The logon server could not validate the logon.
  {$EXTERNALSYM NERR_NonValidatedLogon}

//
//      ACF API related (access, user, group)
//              Error codes from BASE+119 to BASE+149


  NERR_ACFNotFound          = (NERR_BASE+119); // The security database could not be found.
  {$EXTERNALSYM NERR_ACFNotFound}
  NERR_GroupNotFound        = (NERR_BASE+120); // The group name could not be found.
  {$EXTERNALSYM NERR_GroupNotFound}
  NERR_UserNotFound         = (NERR_BASE+121); // The user name could not be found.
  {$EXTERNALSYM NERR_UserNotFound}
  NERR_ResourceNotFound     = (NERR_BASE+122); // The resource name could not be found.
  {$EXTERNALSYM NERR_ResourceNotFound}
  NERR_GroupExists          = (NERR_BASE+123); // The group already exists.
  {$EXTERNALSYM NERR_GroupExists}
  NERR_UserExists           = (NERR_BASE+124); // The account already exists.
  {$EXTERNALSYM NERR_UserExists}
  NERR_ResourceExists       = (NERR_BASE+125); // The resource permission list already exists.
  {$EXTERNALSYM NERR_ResourceExists}
  NERR_NotPrimary           = (NERR_BASE+126); // This operation is only allowed on the primary domain controller of the domain.
  {$EXTERNALSYM NERR_NotPrimary}
  NERR_ACFNotLoaded         = (NERR_BASE+127); // The security database has not been started.
  {$EXTERNALSYM NERR_ACFNotLoaded}
  NERR_ACFNoRoom            = (NERR_BASE+128); // There are too many names in the user accounts database.
  {$EXTERNALSYM NERR_ACFNoRoom}
  NERR_ACFFileIOFail        = (NERR_BASE+129); // A disk I/O failure occurred.
  {$EXTERNALSYM NERR_ACFFileIOFail}
  NERR_ACFTooManyLists      = (NERR_BASE+130); // The limit of 64 entries per resource was exceeded.
  {$EXTERNALSYM NERR_ACFTooManyLists}
  NERR_UserLogon            = (NERR_BASE+131); // Deleting a user with a session is not allowed.
  {$EXTERNALSYM NERR_UserLogon}
  NERR_ACFNoParent          = (NERR_BASE+132); // The parent directory could not be located.
  {$EXTERNALSYM NERR_ACFNoParent}
  NERR_CanNotGrowSegment    = (NERR_BASE+133); // Unable to add to the security database session cache segment.
  {$EXTERNALSYM NERR_CanNotGrowSegment}
  NERR_SpeGroupOp           = (NERR_BASE+134); // This operation is not allowed on this special group.
  {$EXTERNALSYM NERR_SpeGroupOp}
  NERR_NotInCache           = (NERR_BASE+135); // This user is not cached in user accounts database session cache.
  {$EXTERNALSYM NERR_NotInCache}
  NERR_UserInGroup          = (NERR_BASE+136); // The user already belongs to this group.
  {$EXTERNALSYM NERR_UserInGroup}
  NERR_UserNotInGroup       = (NERR_BASE+137); // The user does not belong to this group.
  {$EXTERNALSYM NERR_UserNotInGroup}
  NERR_AccountUndefined     = (NERR_BASE+138); // This user account is undefined.
  {$EXTERNALSYM NERR_AccountUndefined}
  NERR_AccountExpired       = (NERR_BASE+139); // This user account has expired.
  {$EXTERNALSYM NERR_AccountExpired}
  NERR_InvalidWorkstation   = (NERR_BASE+140); // The user is not allowed to log on from this workstation.
  {$EXTERNALSYM NERR_InvalidWorkstation}
  NERR_InvalidLogonHours    = (NERR_BASE+141); // The user is not allowed to log on at this time.
  {$EXTERNALSYM NERR_InvalidLogonHours}
  NERR_PasswordExpired      = (NERR_BASE+142); // The password of this user has expired.
  {$EXTERNALSYM NERR_PasswordExpired}
  NERR_PasswordCantChange   = (NERR_BASE+143); // The password of this user cannot change.
  {$EXTERNALSYM NERR_PasswordCantChange}
  NERR_PasswordHistConflict = (NERR_BASE+144); // This password cannot be used now.
  {$EXTERNALSYM NERR_PasswordHistConflict}
  NERR_PasswordTooShort     = (NERR_BASE+145); // The password does not meet the password policy requirements. Check the minimum password length, password complexity and password history requirements.
  {$EXTERNALSYM NERR_PasswordTooShort}
  NERR_PasswordTooRecent    = (NERR_BASE+146); // The password of this user is too recent to change.
  {$EXTERNALSYM NERR_PasswordTooRecent}
  NERR_InvalidDatabase      = (NERR_BASE+147); // The security database is corrupted.
  {$EXTERNALSYM NERR_InvalidDatabase}
  NERR_DatabaseUpToDate     = (NERR_BASE+148); // No updates are necessary to this replicant network/local security database.
  {$EXTERNALSYM NERR_DatabaseUpToDate}
  NERR_SyncRequired         = (NERR_BASE+149); // This replicant database is outdated; synchronization is required.
  {$EXTERNALSYM NERR_SyncRequired}

//
//      Use API related
//              Error codes from BASE+150 to BASE+169


  NERR_UseNotFound    = (NERR_BASE+150); // The network connection could not be found.
  {$EXTERNALSYM NERR_UseNotFound}
  NERR_BadAsgType     = (NERR_BASE+151); // This asg_type is invalid.
  {$EXTERNALSYM NERR_BadAsgType}
  NERR_DeviceIsShared = (NERR_BASE+152); // This device is currently being shared.
  {$EXTERNALSYM NERR_DeviceIsShared}

//
//      Message Server related
//              Error codes BASE+170 to BASE+209


  NERR_NoComputerName     = (NERR_BASE+170); // The computer name could not be added as a message alias.  The name may already exist on the network.
  {$EXTERNALSYM NERR_NoComputerName}
  NERR_MsgAlreadyStarted  = (NERR_BASE+171); // The Messenger service is already started.
  {$EXTERNALSYM NERR_MsgAlreadyStarted}
  NERR_MsgInitFailed      = (NERR_BASE+172); // The Messenger service failed to start.
  {$EXTERNALSYM NERR_MsgInitFailed}
  NERR_NameNotFound       = (NERR_BASE+173); // The message alias could not be found on the network.
  {$EXTERNALSYM NERR_NameNotFound}
  NERR_AlreadyForwarded   = (NERR_BASE+174); // This message alias has already been forwarded.
  {$EXTERNALSYM NERR_AlreadyForwarded}
  NERR_AddForwarded       = (NERR_BASE+175); // This message alias has been added but is still forwarded.
  {$EXTERNALSYM NERR_AddForwarded}
  NERR_AlreadyExists      = (NERR_BASE+176); // This message alias already exists locally.
  {$EXTERNALSYM NERR_AlreadyExists}
  NERR_TooManyNames       = (NERR_BASE+177); // The maximum number of added message aliases has been exceeded.
  {$EXTERNALSYM NERR_TooManyNames}
  NERR_DelComputerName    = (NERR_BASE+178); // The computer name could not be deleted.
  {$EXTERNALSYM NERR_DelComputerName}
  NERR_LocalForward       = (NERR_BASE+179); // Messages cannot be forwarded back to the same workstation.
  {$EXTERNALSYM NERR_LocalForward}
  NERR_GrpMsgProcessor    = (NERR_BASE+180); // An error occurred in the domain message processor.
  {$EXTERNALSYM NERR_GrpMsgProcessor}
  NERR_PausedRemote       = (NERR_BASE+181); // The message was sent, but the recipient has paused the Messenger service.
  {$EXTERNALSYM NERR_PausedRemote}
  NERR_BadReceive         = (NERR_BASE+182); // The message was sent but not received.
  {$EXTERNALSYM NERR_BadReceive}
  NERR_NameInUse          = (NERR_BASE+183); // The message alias is currently in use. Try again later.
  {$EXTERNALSYM NERR_NameInUse}
  NERR_MsgNotStarted      = (NERR_BASE+184); // The Messenger service has not been started.
  {$EXTERNALSYM NERR_MsgNotStarted}
  NERR_NotLocalName       = (NERR_BASE+185); // The name is not on the local computer.
  {$EXTERNALSYM NERR_NotLocalName}
  NERR_NoForwardName      = (NERR_BASE+186); // The forwarded message alias could not be found on the network.
  {$EXTERNALSYM NERR_NoForwardName}
  NERR_RemoteFull         = (NERR_BASE+187); // The message alias table on the remote station is full.
  {$EXTERNALSYM NERR_RemoteFull}
  NERR_NameNotForwarded   = (NERR_BASE+188); // Messages for this alias are not currently being forwarded.
  {$EXTERNALSYM NERR_NameNotForwarded}
  NERR_TruncatedBroadcast = (NERR_BASE+189); // The broadcast message was truncated.
  {$EXTERNALSYM NERR_TruncatedBroadcast}
  NERR_InvalidDevice      = (NERR_BASE+194); // This is an invalid device name.
  {$EXTERNALSYM NERR_InvalidDevice}
  NERR_WriteFault         = (NERR_BASE+195); // A write fault occurred.
  {$EXTERNALSYM NERR_WriteFault}
// UNUSED BASE+196
  NERR_DuplicateName = (NERR_BASE+197); // A duplicate message alias exists on the network.
  {$EXTERNALSYM NERR_DuplicateName}
  NERR_DeleteLater   = (NERR_BASE+198); // @W This message alias will be deleted later.
  {$EXTERNALSYM NERR_DeleteLater}
  NERR_IncompleteDel = (NERR_BASE+199); // The message alias was not successfully deleted from all networks.
  {$EXTERNALSYM NERR_IncompleteDel}
  NERR_MultipleNets  = (NERR_BASE+200); // This operation is not supported on computers with multiple networks.
  {$EXTERNALSYM NERR_MultipleNets}

//
//      Server API related
//             Error codes BASE+210 to BASE+229


  NERR_NetNameNotFound        = (NERR_BASE+210); // This shared resource does not exist.
  {$EXTERNALSYM NERR_NetNameNotFound}
  NERR_DeviceNotShared        = (NERR_BASE+211); // This device is not shared.
  {$EXTERNALSYM NERR_DeviceNotShared}
  NERR_ClientNameNotFound     = (NERR_BASE+212); // A session does not exist with that computer name.
  {$EXTERNALSYM NERR_ClientNameNotFound}
  NERR_FileIdNotFound         = (NERR_BASE+214); // There is not an open file with that identification number.
  {$EXTERNALSYM NERR_FileIdNotFound}
  NERR_ExecFailure            = (NERR_BASE+215); // A failure occurred when executing a remote administration command.
  {$EXTERNALSYM NERR_ExecFailure}
  NERR_TmpFile                = (NERR_BASE+216); // A failure occurred when opening a remote temporary file.
  {$EXTERNALSYM NERR_TmpFile}
  NERR_TooMuchData            = (NERR_BASE+217); // The data returned from a remote administration command has been truncated to 64K.
  {$EXTERNALSYM NERR_TooMuchData}
  NERR_DeviceShareConflict    = (NERR_BASE+218); // This device cannot be shared as both a spooled and a non-spooled resource.
  {$EXTERNALSYM NERR_DeviceShareConflict}
  NERR_BrowserTableIncomplete = (NERR_BASE+219); // The information in the list of servers may be incorrect.
  {$EXTERNALSYM NERR_BrowserTableIncomplete}
  NERR_NotLocalDomain         = (NERR_BASE+220); // The computer is not active in this domain.
  {$EXTERNALSYM NERR_NotLocalDomain}
  NERR_IsDfsShare             = (NERR_BASE+221); // The share must be removed from the Distributed File System before it can be deleted.
  {$EXTERNALSYM NERR_IsDfsShare}

//
//      CharDev API related
//              Error codes BASE+230 to BASE+249


// UNUSED BASE+230
  NERR_DevInvalidOpCode  = (NERR_BASE+231); // The operation is invalid for this device.
  {$EXTERNALSYM NERR_DevInvalidOpCode}
  NERR_DevNotFound       = (NERR_BASE+232); // This device cannot be shared.
  {$EXTERNALSYM NERR_DevNotFound}
  NERR_DevNotOpen        = (NERR_BASE+233); // This device was not open.
  {$EXTERNALSYM NERR_DevNotOpen}
  NERR_BadQueueDevString = (NERR_BASE+234); // This device name list is invalid.
  {$EXTERNALSYM NERR_BadQueueDevString}
  NERR_BadQueuePriority  = (NERR_BASE+235); // The queue priority is invalid.
  {$EXTERNALSYM NERR_BadQueuePriority}
  NERR_NoCommDevs        = (NERR_BASE+237); // There are no shared communication devices.
  {$EXTERNALSYM NERR_NoCommDevs}
  NERR_QueueNotFound     = (NERR_BASE+238); // The queue you specified does not exist.
  {$EXTERNALSYM NERR_QueueNotFound}
  NERR_BadDevString      = (NERR_BASE+240); // This list of devices is invalid.
  {$EXTERNALSYM NERR_BadDevString}
  NERR_BadDev            = (NERR_BASE+241); // The requested device is invalid.
  {$EXTERNALSYM NERR_BadDev}
  NERR_InUseBySpooler    = (NERR_BASE+242); // This device is already in use by the spooler.
  {$EXTERNALSYM NERR_InUseBySpooler}
  NERR_CommDevInUse      = (NERR_BASE+243); // This device is already in use as a communication device.
  {$EXTERNALSYM NERR_CommDevInUse}

//
//      NetICanonicalize and NetIType and NetIMakeLMFileName
//      NetIListCanon and NetINameCheck
//              Error codes BASE+250 to BASE+269


  NERR_InvalidComputer = (NERR_BASE+251); // This computer name is invalid.
  {$EXTERNALSYM NERR_InvalidComputer}
// UNUSED BASE+252
// UNUSED BASE+253
  NERR_MaxLenExceeded = (NERR_BASE+254); // The string and prefix specified are too long.
  {$EXTERNALSYM NERR_MaxLenExceeded}
// UNUSED BASE+255
  NERR_BadComponent = (NERR_BASE+256); // This path component is invalid.
  {$EXTERNALSYM NERR_BadComponent}
  NERR_CantType     = (NERR_BASE+257); // Could not determine the type of input.
  {$EXTERNALSYM NERR_CantType}
// UNUSED BASE+258
// UNUSED BASE+259
  NERR_TooManyEntries = (NERR_BASE+262); // The buffer for types is not big enough.
  {$EXTERNALSYM NERR_TooManyEntries}

//
//      NetProfile
//              Error codes BASE+270 to BASE+276


  NERR_ProfileFileTooBig = (NERR_BASE+270); // Profile files cannot exceed 64K.
  {$EXTERNALSYM NERR_ProfileFileTooBig}
  NERR_ProfileOffset     = (NERR_BASE+271); // The start offset is out of range.
  {$EXTERNALSYM NERR_ProfileOffset}
  NERR_ProfileCleanup    = (NERR_BASE+272); // The system cannot delete current connections to network resources.
  {$EXTERNALSYM NERR_ProfileCleanup}
  NERR_ProfileUnknownCmd = (NERR_BASE+273); // The system was unable to parse the command line in this file.
  {$EXTERNALSYM NERR_ProfileUnknownCmd}
  NERR_ProfileLoadErr    = (NERR_BASE+274); // An error occurred while loading the profile file.
  {$EXTERNALSYM NERR_ProfileLoadErr}
  NERR_ProfileSaveErr    = (NERR_BASE+275); // @W Errors occurred while saving the profile file.  The profile was partially saved.
  {$EXTERNALSYM NERR_ProfileSaveErr}


//
//      NetAudit and NetErrorLog
//              Error codes BASE+277 to BASE+279


  NERR_LogOverflow    = (NERR_BASE+277); // Log file %1 is full.
  {$EXTERNALSYM NERR_LogOverflow}
  NERR_LogFileChanged = (NERR_BASE+278); // This log file has changed between reads.
  {$EXTERNALSYM NERR_LogFileChanged}
  NERR_LogFileCorrupt = (NERR_BASE+279); // Log file %1 is corrupt.
  {$EXTERNALSYM NERR_LogFileCorrupt}


//
//      NetRemote
//              Error codes BASE+280 to BASE+299

  NERR_SourceIsDir      = (NERR_BASE+280); // The source path cannot be a directory.
  {$EXTERNALSYM NERR_SourceIsDir}
  NERR_BadSource        = (NERR_BASE+281); // The source path is illegal.
  {$EXTERNALSYM NERR_BadSource}
  NERR_BadDest          = (NERR_BASE+282); // The destination path is illegal.
  {$EXTERNALSYM NERR_BadDest}
  NERR_DifferentServers = (NERR_BASE+283); // The source and destination paths are on different servers.
  {$EXTERNALSYM NERR_DifferentServers}
// UNUSED BASE+284
  NERR_RunSrvPaused = (NERR_BASE+285); // The Run server you requested is paused.
  {$EXTERNALSYM NERR_RunSrvPaused}
// UNUSED BASE+286
// UNUSED BASE+287
// UNUSED BASE+288
  NERR_ErrCommRunSrv = (NERR_BASE+289); // An error occurred when communicating with a Run server.
  {$EXTERNALSYM NERR_ErrCommRunSrv}
// UNUSED BASE+290
  NERR_ErrorExecingGhost = (NERR_BASE+291); // An error occurred when starting a background process.
  {$EXTERNALSYM NERR_ErrorExecingGhost}
  NERR_ShareNotFound     = (NERR_BASE+292); // The shared resource you are connected to could not be found.
  {$EXTERNALSYM NERR_ShareNotFound}
// UNUSED BASE+293
// UNUSED BASE+294


//
//  NetWksta.sys (redir) returned error codes.
//
//          NERR_BASE + (300-329)


  NERR_InvalidLana     = (NERR_BASE+300); // The LAN adapter number is invalid.
  {$EXTERNALSYM NERR_InvalidLana}
  NERR_OpenFiles       = (NERR_BASE+301); // There are open files on the connection.
  {$EXTERNALSYM NERR_OpenFiles}
  NERR_ActiveConns     = (NERR_BASE+302); // Active connections still exist.
  {$EXTERNALSYM NERR_ActiveConns}
  NERR_BadPasswordCore = (NERR_BASE+303); // This share name or password is invalid.
  {$EXTERNALSYM NERR_BadPasswordCore}
  NERR_DevInUse        = (NERR_BASE+304); // The device is being accessed by an active process.
  {$EXTERNALSYM NERR_DevInUse}
  NERR_LocalDrive      = (NERR_BASE+305); // The drive letter is in use locally.
  {$EXTERNALSYM NERR_LocalDrive}

//
//  Alert error codes.
//
//          NERR_BASE + (330-339)

  NERR_AlertExists       = (NERR_BASE+330); // The specified client is already registered for the specified event.
  {$EXTERNALSYM NERR_AlertExists}
  NERR_TooManyAlerts     = (NERR_BASE+331); // The alert table is full.
  {$EXTERNALSYM NERR_TooManyAlerts}
  NERR_NoSuchAlert       = (NERR_BASE+332); // An invalid or nonexistent alert name was raised.
  {$EXTERNALSYM NERR_NoSuchAlert}
  NERR_BadRecipient      = (NERR_BASE+333); // The alert recipient is invalid.
  {$EXTERNALSYM NERR_BadRecipient}
  NERR_AcctLimitExceeded = (NERR_BASE+334); // A user's session with this server has been deleted
  {$EXTERNALSYM NERR_AcctLimitExceeded}
                                                // because the user's logon hours are no longer valid.

//
//  Additional Error and Audit log codes.
//
//          NERR_BASE +(340-343)

  NERR_InvalidLogSeek = (NERR_BASE+340); // The log file does not contain the requested record number.
  {$EXTERNALSYM NERR_InvalidLogSeek}
// UNUSED BASE+341
// UNUSED BASE+342
// UNUSED BASE+343

//
//  Additional UAS and NETLOGON codes
//
//          NERR_BASE +(350-359)

  NERR_BadUasConfig       = (NERR_BASE+350); // The user accounts database is not configured correctly.
  {$EXTERNALSYM NERR_BadUasConfig}
  NERR_InvalidUASOp       = (NERR_BASE+351); // This operation is not permitted when the Netlogon service is running.
  {$EXTERNALSYM NERR_InvalidUASOp}
  NERR_LastAdmin          = (NERR_BASE+352); // This operation is not allowed on the last administrative account.
  {$EXTERNALSYM NERR_LastAdmin}
  NERR_DCNotFound         = (NERR_BASE+353); // Could not find domain controller for this domain.
  {$EXTERNALSYM NERR_DCNotFound}
  NERR_LogonTrackingError = (NERR_BASE+354); // Could not set logon information for this user.
  {$EXTERNALSYM NERR_LogonTrackingError}
  NERR_NetlogonNotStarted = (NERR_BASE+355); // The Netlogon service has not been started.
  {$EXTERNALSYM NERR_NetlogonNotStarted}
  NERR_CanNotGrowUASFile  = (NERR_BASE+356); // Unable to add to the user accounts database.
  {$EXTERNALSYM NERR_CanNotGrowUASFile}
  NERR_TimeDiffAtDC       = (NERR_BASE+357); // This server's clock is not synchronized with the primary domain controller's clock.
  {$EXTERNALSYM NERR_TimeDiffAtDC}
  NERR_PasswordMismatch   = (NERR_BASE+358); // A password mismatch has been detected.
  {$EXTERNALSYM NERR_PasswordMismatch}


//
//  Server Integration error codes.
//
//          NERR_BASE +(360-369)

  NERR_NoSuchServer       = (NERR_BASE+360); // The server identification does not specify a valid server.
  {$EXTERNALSYM NERR_NoSuchServer}
  NERR_NoSuchSession      = (NERR_BASE+361); // The session identification does not specify a valid session.
  {$EXTERNALSYM NERR_NoSuchSession}
  NERR_NoSuchConnection   = (NERR_BASE+362); // The connection identification does not specify a valid connection.
  {$EXTERNALSYM NERR_NoSuchConnection}
  NERR_TooManyServers     = (NERR_BASE+363); // There is no space for another entry in the table of available servers.
  {$EXTERNALSYM NERR_TooManyServers}
  NERR_TooManySessions    = (NERR_BASE+364); // The server has reached the maximum number of sessions it supports.
  {$EXTERNALSYM NERR_TooManySessions}
  NERR_TooManyConnections = (NERR_BASE+365); // The server has reached the maximum number of connections it supports.
  {$EXTERNALSYM NERR_TooManyConnections}
  NERR_TooManyFiles       = (NERR_BASE+366); // The server cannot open more files because it has reached its maximum number.
  {$EXTERNALSYM NERR_TooManyFiles}
  NERR_NoAlternateServers = (NERR_BASE+367); // There are no alternate servers registered on this server.
  {$EXTERNALSYM NERR_NoAlternateServers}
// UNUSED BASE+368
// UNUSED BASE+369

  NERR_TryDownLevel = (NERR_BASE+370); // Try down-level (remote admin protocol) version of API instead.
  {$EXTERNALSYM NERR_TryDownLevel}

//
//  UPS error codes.
//
//          NERR_BASE + (380-384)

  NERR_UPSDriverNotStarted = (NERR_BASE+380); // The UPS driver could not be accessed by the UPS service.
  {$EXTERNALSYM NERR_UPSDriverNotStarted}
  NERR_UPSInvalidConfig    = (NERR_BASE+381); // The UPS service is not configured correctly.
  {$EXTERNALSYM NERR_UPSInvalidConfig}
  NERR_UPSInvalidCommPort  = (NERR_BASE+382); // The UPS service could not access the specified Comm Port.
  {$EXTERNALSYM NERR_UPSInvalidCommPort}
  NERR_UPSSignalAsserted   = (NERR_BASE+383); // The UPS indicated a line fail or low battery situation. Service not started.
  {$EXTERNALSYM NERR_UPSSignalAsserted}
  NERR_UPSShutdownFailed   = (NERR_BASE+384); // The UPS service failed to perform a system shut down.
  {$EXTERNALSYM NERR_UPSShutdownFailed}

//
//  Remoteboot error codes.
//
//           NERR_BASE + (400-419)
//           Error codes 400 - 405 are used by RPLBOOT.SYS.
//           Error codes 403, 407 - 416 are used by RPLLOADR.COM,
//           Error code 417 is the alerter message of REMOTEBOOT (RPLSERVR.EXE).
//           Error code 418 is for when REMOTEBOOT can't start
//           Error code 419 is for a disallowed 2nd rpl connection
//

  NERR_BadDosRetCode      = (NERR_BASE+400); // The program below returned an MS-DOS error code:
  {$EXTERNALSYM NERR_BadDosRetCode}
  NERR_ProgNeedsExtraMem  = (NERR_BASE+401); // The program below needs more memory:
  {$EXTERNALSYM NERR_ProgNeedsExtraMem}
  NERR_BadDosFunction     = (NERR_BASE+402); // The program below called an unsupported MS-DOS function:
  {$EXTERNALSYM NERR_BadDosFunction}
  NERR_RemoteBootFailed   = (NERR_BASE+403); // The workstation failed to boot.
  {$EXTERNALSYM NERR_RemoteBootFailed}
  NERR_BadFileCheckSum    = (NERR_BASE+404); // The file below is corrupt.
  {$EXTERNALSYM NERR_BadFileCheckSum}
  NERR_NoRplBootSystem    = (NERR_BASE+405); // No loader is specified in the boot-block definition file.
  {$EXTERNALSYM NERR_NoRplBootSystem}
  NERR_RplLoadrNetBiosErr = (NERR_BASE+406); // NetBIOS returned an error: The NCB and SMB are dumped above.
  {$EXTERNALSYM NERR_RplLoadrNetBiosErr}
  NERR_RplLoadrDiskErr    = (NERR_BASE+407); // A disk I/O error occurred.
  {$EXTERNALSYM NERR_RplLoadrDiskErr}
  NERR_ImageParamErr      = (NERR_BASE+408); // Image parameter substitution failed.
  {$EXTERNALSYM NERR_ImageParamErr}
  NERR_TooManyImageParams = (NERR_BASE+409); // Too many image parameters cross disk sector boundaries.
  {$EXTERNALSYM NERR_TooManyImageParams}
  NERR_NonDosFloppyUsed   = (NERR_BASE+410); // The image was not generated from an MS-DOS diskette formatted with /S.
  {$EXTERNALSYM NERR_NonDosFloppyUsed}
  NERR_RplBootRestart     = (NERR_BASE+411); // Remote boot will be restarted later.
  {$EXTERNALSYM NERR_RplBootRestart}
  NERR_RplSrvrCallFailed  = (NERR_BASE+412); // The call to the Remoteboot server failed.
  {$EXTERNALSYM NERR_RplSrvrCallFailed}
  NERR_CantConnectRplSrvr = (NERR_BASE+413); // Cannot connect to the Remoteboot server.
  {$EXTERNALSYM NERR_CantConnectRplSrvr}
  NERR_CantOpenImageFile  = (NERR_BASE+414); // Cannot open image file on the Remoteboot server.
  {$EXTERNALSYM NERR_CantOpenImageFile}
  NERR_CallingRplSrvr     = (NERR_BASE+415); // Connecting to the Remoteboot server...
  {$EXTERNALSYM NERR_CallingRplSrvr}
  NERR_StartingRplBoot    = (NERR_BASE+416); // Connecting to the Remoteboot server...
  {$EXTERNALSYM NERR_StartingRplBoot}
  NERR_RplBootServiceTerm = (NERR_BASE+417); // Remote boot service was stopped; check the error log for the cause of the problem.
  {$EXTERNALSYM NERR_RplBootServiceTerm}
  NERR_RplBootStartFailed = (NERR_BASE+418); // Remote boot startup failed; check the error log for the cause of the problem.
  {$EXTERNALSYM NERR_RplBootStartFailed}
  NERR_RPL_CONNECTED      = (NERR_BASE+419); // A second connection to a Remoteboot resource is not allowed.
  {$EXTERNALSYM NERR_RPL_CONNECTED}

//
//  FTADMIN API error codes
//
//       NERR_BASE + (425-434)
//
//       (Currently not used in NT)
//


//
//  Browser service API error codes
//
//       NERR_BASE + (450-475)
//

  NERR_BrowserConfiguredToNotRun = (NERR_BASE+450); // The browser service was configured with MaintainServerList=No.
  {$EXTERNALSYM NERR_BrowserConfiguredToNotRun}

//
//  Additional Remoteboot error codes.
//
//          NERR_BASE + (510-550)

  NERR_RplNoAdaptersStarted      = (NERR_BASE+510); // Service failed to start since none of the network adapters started with this service.
  {$EXTERNALSYM NERR_RplNoAdaptersStarted}
  NERR_RplBadRegistry            = (NERR_BASE+511); // Service failed to start due to bad startup information in the registry.
  {$EXTERNALSYM NERR_RplBadRegistry}
  NERR_RplBadDatabase            = (NERR_BASE+512); // Service failed to start because its database is absent or corrupt.
  {$EXTERNALSYM NERR_RplBadDatabase}
  NERR_RplRplfilesShare          = (NERR_BASE+513); // Service failed to start because RPLFILES share is absent.
  {$EXTERNALSYM NERR_RplRplfilesShare}
  NERR_RplNotRplServer           = (NERR_BASE+514); // Service failed to start because RPLUSER group is absent.
  {$EXTERNALSYM NERR_RplNotRplServer}
  NERR_RplCannotEnum             = (NERR_BASE+515); // Cannot enumerate service records.
  {$EXTERNALSYM NERR_RplCannotEnum}
  NERR_RplWkstaInfoCorrupted     = (NERR_BASE+516); // Workstation record information has been corrupted.
  {$EXTERNALSYM NERR_RplWkstaInfoCorrupted}
  NERR_RplWkstaNotFound          = (NERR_BASE+517); // Workstation record was not found.
  {$EXTERNALSYM NERR_RplWkstaNotFound}
  NERR_RplWkstaNameUnavailable   = (NERR_BASE+518); // Workstation name is in use by some other workstation.
  {$EXTERNALSYM NERR_RplWkstaNameUnavailable}
  NERR_RplProfileInfoCorrupted   = (NERR_BASE+519); // Profile record information has been corrupted.
  {$EXTERNALSYM NERR_RplProfileInfoCorrupted}
  NERR_RplProfileNotFound        = (NERR_BASE+520); // Profile record was not found.
  {$EXTERNALSYM NERR_RplProfileNotFound}
  NERR_RplProfileNameUnavailable = (NERR_BASE+521); // Profile name is in use by some other profile.
  {$EXTERNALSYM NERR_RplProfileNameUnavailable}
  NERR_RplProfileNotEmpty        = (NERR_BASE+522); // There are workstations using this profile.
  {$EXTERNALSYM NERR_RplProfileNotEmpty}
  NERR_RplConfigInfoCorrupted    = (NERR_BASE+523); // Configuration record information has been corrupted.
  {$EXTERNALSYM NERR_RplConfigInfoCorrupted}
  NERR_RplConfigNotFound         = (NERR_BASE+524); // Configuration record was not found.
  {$EXTERNALSYM NERR_RplConfigNotFound}
  NERR_RplAdapterInfoCorrupted   = (NERR_BASE+525); // Adapter id record information has been corrupted.
  {$EXTERNALSYM NERR_RplAdapterInfoCorrupted}
  NERR_RplInternal               = (NERR_BASE+526); // An internal service error has occurred.
  {$EXTERNALSYM NERR_RplInternal}
  NERR_RplVendorInfoCorrupted    = (NERR_BASE+527); // Vendor id record information has been corrupted.
  {$EXTERNALSYM NERR_RplVendorInfoCorrupted}
  NERR_RplBootInfoCorrupted      = (NERR_BASE+528); // Boot block record information has been corrupted.
  {$EXTERNALSYM NERR_RplBootInfoCorrupted}
  NERR_RplWkstaNeedsUserAcct     = (NERR_BASE+529); // The user account for this workstation record is missing.
  {$EXTERNALSYM NERR_RplWkstaNeedsUserAcct}
  NERR_RplNeedsRPLUSERAcct       = (NERR_BASE+530); // The RPLUSER local group could not be found.
  {$EXTERNALSYM NERR_RplNeedsRPLUSERAcct}
  NERR_RplBootNotFound           = (NERR_BASE+531); // Boot block record was not found.
  {$EXTERNALSYM NERR_RplBootNotFound}
  NERR_RplIncompatibleProfile    = (NERR_BASE+532); // Chosen profile is incompatible with this workstation.
  {$EXTERNALSYM NERR_RplIncompatibleProfile}
  NERR_RplAdapterNameUnavailable = (NERR_BASE+533); // Chosen network adapter id is in use by some other workstation.
  {$EXTERNALSYM NERR_RplAdapterNameUnavailable}
  NERR_RplConfigNotEmpty         = (NERR_BASE+534); // There are profiles using this configuration.
  {$EXTERNALSYM NERR_RplConfigNotEmpty}
  NERR_RplBootInUse              = (NERR_BASE+535); // There are workstations, profiles or configurations using this boot block.
  {$EXTERNALSYM NERR_RplBootInUse}
  NERR_RplBackupDatabase         = (NERR_BASE+536); // Service failed to backup Remoteboot database.
  {$EXTERNALSYM NERR_RplBackupDatabase}
  NERR_RplAdapterNotFound        = (NERR_BASE+537); // Adapter record was not found.
  {$EXTERNALSYM NERR_RplAdapterNotFound}
  NERR_RplVendorNotFound         = (NERR_BASE+538); // Vendor record was not found.
  {$EXTERNALSYM NERR_RplVendorNotFound}
  NERR_RplVendorNameUnavailable  = (NERR_BASE+539); // Vendor name is in use by some other vendor record.
  {$EXTERNALSYM NERR_RplVendorNameUnavailable}
  NERR_RplBootNameUnavailable    = (NERR_BASE+540); // (boot name, vendor id) is in use by some other boot block record.
  {$EXTERNALSYM NERR_RplBootNameUnavailable}
  NERR_RplConfigNameUnavailable  = (NERR_BASE+541); // Configuration name is in use by some other configuration.
  {$EXTERNALSYM NERR_RplConfigNameUnavailable}

//*INTERNAL_ONLY*

//
//  Dfs API error codes.
//
//          NERR_BASE + (560-590)


  NERR_DfsInternalCorruption        = (NERR_BASE+560); // The internal database maintained by the DFS service is corrupt
  {$EXTERNALSYM NERR_DfsInternalCorruption}
  NERR_DfsVolumeDataCorrupt         = (NERR_BASE+561); // One of the records in the internal DFS database is corrupt
  {$EXTERNALSYM NERR_DfsVolumeDataCorrupt}
  NERR_DfsNoSuchVolume              = (NERR_BASE+562); // There is no DFS name whose entry path matches the input Entry Path
  {$EXTERNALSYM NERR_DfsNoSuchVolume}
  NERR_DfsVolumeAlreadyExists       = (NERR_BASE+563); // A root or link with the given name already exists
  {$EXTERNALSYM NERR_DfsVolumeAlreadyExists}
  NERR_DfsAlreadyShared             = (NERR_BASE+564); // The server share specified is already shared in the DFS
  {$EXTERNALSYM NERR_DfsAlreadyShared}
  NERR_DfsNoSuchShare               = (NERR_BASE+565); // The indicated server share does not support the indicated DFS namespace
  {$EXTERNALSYM NERR_DfsNoSuchShare}
  NERR_DfsNotALeafVolume            = (NERR_BASE+566); // The operation is not valid on this portion of the namespace
  {$EXTERNALSYM NERR_DfsNotALeafVolume}
  NERR_DfsLeafVolume                = (NERR_BASE+567); // The operation is not valid on this portion of the namespace
  {$EXTERNALSYM NERR_DfsLeafVolume}
  NERR_DfsVolumeHasMultipleServers  = (NERR_BASE+568); // The operation is ambiguous because the link has multiple servers
  {$EXTERNALSYM NERR_DfsVolumeHasMultipleServers}
  NERR_DfsCantCreateJunctionPoint   = (NERR_BASE+569); // Unable to create a link
  {$EXTERNALSYM NERR_DfsCantCreateJunctionPoint}
  NERR_DfsServerNotDfsAware         = (NERR_BASE+570); // The server is not DFS Aware
  {$EXTERNALSYM NERR_DfsServerNotDfsAware}
  NERR_DfsBadRenamePath             = (NERR_BASE+571); // The specified rename target path is invalid
  {$EXTERNALSYM NERR_DfsBadRenamePath}
  NERR_DfsVolumeIsOffline           = (NERR_BASE+572); // The specified DFS link is offline
  {$EXTERNALSYM NERR_DfsVolumeIsOffline}
  NERR_DfsNoSuchServer              = (NERR_BASE+573); // The specified server is not a server for this link
  {$EXTERNALSYM NERR_DfsNoSuchServer}
  NERR_DfsCyclicalName              = (NERR_BASE+574); // A cycle in the DFS name was detected
  {$EXTERNALSYM NERR_DfsCyclicalName}
  NERR_DfsNotSupportedInServerDfs   = (NERR_BASE+575); // The operation is not supported on a server-based DFS
  {$EXTERNALSYM NERR_DfsNotSupportedInServerDfs}
  NERR_DfsDuplicateService          = (NERR_BASE+576); // This link is already supported by the specified server-share
  {$EXTERNALSYM NERR_DfsDuplicateService}
  NERR_DfsCantRemoveLastServerShare = (NERR_BASE+577); // Can't remove the last server-share supporting this root or link
  {$EXTERNALSYM NERR_DfsCantRemoveLastServerShare}
  NERR_DfsVolumeIsInterDfs          = (NERR_BASE+578); // The operation is not supported for an Inter-DFS link
  {$EXTERNALSYM NERR_DfsVolumeIsInterDfs}
  NERR_DfsInconsistent              = (NERR_BASE+579); // The internal state of the DFS Service has become inconsistent
  {$EXTERNALSYM NERR_DfsInconsistent}
  NERR_DfsServerUpgraded            = (NERR_BASE+580); // The DFS Service has been installed on the specified server
  {$EXTERNALSYM NERR_DfsServerUpgraded}
  NERR_DfsDataIsIdentical           = (NERR_BASE+581); // The DFS data being reconciled is identical
  {$EXTERNALSYM NERR_DfsDataIsIdentical}
  NERR_DfsCantRemoveDfsRoot         = (NERR_BASE+582); // The DFS root cannot be deleted - Uninstall DFS if required
  {$EXTERNALSYM NERR_DfsCantRemoveDfsRoot}
  NERR_DfsChildOrParentInDfs        = (NERR_BASE+583); // A child or parent directory of the share is already in a DFS
  {$EXTERNALSYM NERR_DfsChildOrParentInDfs}
  NERR_DfsInternalError             = (NERR_BASE+590); // DFS internal error
  {$EXTERNALSYM NERR_DfsInternalError}

//
//  Net setup error codes.
//
//          NERR_BASE + (591-600)

  NERR_SetupAlreadyJoined           = (NERR_BASE+591); // This machine is already joined to a domain.
  {$EXTERNALSYM NERR_SetupAlreadyJoined}
  NERR_SetupNotJoined               = (NERR_BASE+592); // This machine is not currently joined to a domain.
  {$EXTERNALSYM NERR_SetupNotJoined}
  NERR_SetupDomainController        = (NERR_BASE+593); // This machine is a domain controller and cannot be unjoined from a domain.
  {$EXTERNALSYM NERR_SetupDomainController}
  NERR_DefaultJoinRequired          = (NERR_BASE+594); // The destination domain controller does not support creating machine accounts in OUs.
  {$EXTERNALSYM NERR_DefaultJoinRequired}
  NERR_InvalidWorkgroupName         = (NERR_BASE+595); // The specified workgroup name is invalid.
  {$EXTERNALSYM NERR_InvalidWorkgroupName}
  NERR_NameUsesIncompatibleCodePage = (NERR_BASE+596); // The specified computer name is incompatible with the default language used on the domain controller.
  {$EXTERNALSYM NERR_NameUsesIncompatibleCodePage}
  NERR_ComputerAccountNotFound      = (NERR_BASE+597); // The specified computer account could not be found.
  {$EXTERNALSYM NERR_ComputerAccountNotFound}
  NERR_PersonalSku                  = (NERR_BASE+598); // This version of Windows cannot be joined to a domain.
  {$EXTERNALSYM NERR_PersonalSku}

//
//  Some Password and account error results
//
//          NERR_BASE + (601 - 608)
//

  NERR_PasswordMustChange           = (NERR_BASE + 601);   // Password must change at next logon
  {$EXTERNALSYM NERR_PasswordMustChange}
  NERR_AccountLockedOut             = (NERR_BASE + 602);   // Account is locked out
  {$EXTERNALSYM NERR_AccountLockedOut}
  NERR_PasswordTooLong              = (NERR_BASE + 603);   // Password is too long
  {$EXTERNALSYM NERR_PasswordTooLong}
  NERR_PasswordNotComplexEnough     = (NERR_BASE + 604);   // Password doesn't meet the complexity policy
  {$EXTERNALSYM NERR_PasswordNotComplexEnough}
  NERR_PasswordFilterError          = (NERR_BASE + 605);   // Password doesn't meet the requirements of the filter dll's
  {$EXTERNALSYM NERR_PasswordFilterError}

//**********WARNING ****************
//The range 2750-2799 has been     *
//allocated to the IBM LAN Server  *
//*********************************

//**********WARNING ****************
//The range 2900-2999 has been     *
//reserved for Microsoft OEMs      *
//*********************************

//*END_INTERNAL*

  MAX_NERR = (NERR_BASE+899); // This is the last error in NERR range.
  {$EXTERNALSYM MAX_NERR}

//
// end of list
//
//    WARNING:  Do not exceed MAX_NERR; values above this are used by
//              other error code ranges (errlog.h, service.h, apperr.h).

// JwaLmCons, complete
// LAN Manager common definitions

const
  NetApi32 = 'netapi32.dll';

//
// NOTE:  Lengths of strings are given as the maximum lengths of the
// string in characters (not bytes).  This does not include space for the
// terminating 0-characters.  When allocating space for such an item,
// use the form:
//
//     TCHAR username[UNLEN+1];
//
// Definitions of the form LN20_* define those values in effect for
// LanMan 2.0.
//

//
// String Lengths for various LanMan names
//

const
  CNLEN      = 15; // Computer name length
  {$EXTERNALSYM CNLEN}
  LM20_CNLEN = 15; // LM 2.0 Computer name length
  {$EXTERNALSYM LM20_CNLEN}
  DNLEN      = CNLEN; // Maximum domain name length
  {$EXTERNALSYM DNLEN}
  LM20_DNLEN = LM20_CNLEN; // LM 2.0 Maximum domain name length
  {$EXTERNALSYM LM20_DNLEN}

//#if (CNLEN != DNLEN)
//#error CNLEN and DNLEN are not equal
//#endif

  UNCLEN      = (CNLEN+2); // UNC computer name length
  {$EXTERNALSYM UNCLEN}
  LM20_UNCLEN = (LM20_CNLEN+2); // LM 2.0 UNC computer name length
  {$EXTERNALSYM LM20_UNCLEN}

  NNLEN      = 80; // Net name length (share name)
  {$EXTERNALSYM NNLEN}
  LM20_NNLEN = 12; // LM 2.0 Net name length
  {$EXTERNALSYM LM20_NNLEN}

  RMLEN      = (UNCLEN+1+NNLEN); // Max remote name length
  {$EXTERNALSYM RMLEN}
  LM20_RMLEN = (LM20_UNCLEN+1+LM20_NNLEN); // LM 2.0 Max remote name length
  {$EXTERNALSYM LM20_RMLEN}

  SNLEN        = 80; // Service name length
  {$EXTERNALSYM SNLEN}
  LM20_SNLEN   = 15; // LM 2.0 Service name length
  {$EXTERNALSYM LM20_SNLEN}
  STXTLEN      = 256; // Service text length
  {$EXTERNALSYM STXTLEN}
  LM20_STXTLEN = 63; // LM 2.0 Service text length
  {$EXTERNALSYM LM20_STXTLEN}

  PATHLEN      = 256; // Max. path (not including drive name)
  {$EXTERNALSYM PATHLEN}
  LM20_PATHLEN = 256; // LM 2.0 Max. path
  {$EXTERNALSYM LM20_PATHLEN}

  DEVLEN      = 80; // Device name length
  {$EXTERNALSYM DEVLEN}
  LM20_DEVLEN = 8; // LM 2.0 Device name length
  {$EXTERNALSYM LM20_DEVLEN}

  EVLEN = 16; // Event name length
  {$EXTERNALSYM EVLEN}

//
// User, Group and Password lengths
//

  UNLEN      = 256; // Maximum user name length
  {$EXTERNALSYM UNLEN}
  LM20_UNLEN = 20; // LM 2.0 Maximum user name length
  {$EXTERNALSYM LM20_UNLEN}

  GNLEN      = UNLEN; // Group name
  {$EXTERNALSYM GNLEN}
  LM20_GNLEN = LM20_UNLEN; // LM 2.0 Group name
  {$EXTERNALSYM LM20_GNLEN}

  PWLEN      = 256; // Maximum password length
  {$EXTERNALSYM PWLEN}
  LM20_PWLEN = 14; // LM 2.0 Maximum password length
  {$EXTERNALSYM LM20_PWLEN}

  SHPWLEN = 8; // Share password length (bytes)
  {$EXTERNALSYM SHPWLEN}

  CLTYPE_LEN = 12; // Length of client type string
  {$EXTERNALSYM CLTYPE_LEN}

  MAXCOMMENTSZ      = 256; // Multipurpose comment length
  {$EXTERNALSYM MAXCOMMENTSZ}
  LM20_MAXCOMMENTSZ = 48; // LM 2.0 Multipurpose comment length
  {$EXTERNALSYM LM20_MAXCOMMENTSZ}

  QNLEN      = NNLEN; // Queue name maximum length
  {$EXTERNALSYM QNLEN}
  LM20_QNLEN = LM20_NNLEN; // LM 2.0 Queue name maximum length
  {$EXTERNALSYM LM20_QNLEN}

//#if (QNLEN != NNLEN)
//# error QNLEN and NNLEN are not equal
//#endif

//
// The ALERTSZ and MAXDEVENTRIES defines have not yet been NT'ized.
// Whoever ports these components should change these values appropriately.
//

  ALERTSZ       = 128; // size of alert string in server
  {$EXTERNALSYM ALERTSZ}
  MAXDEVENTRIES = (SizeOf(Integer)*8); // Max number of device entries
  {$EXTERNALSYM MAXDEVENTRIES}

                                        //
                                        // We use int bitmap to represent
                                        //

  NETBIOS_NAME_LEN = 16; // NetBIOS net name (bytes)
  {$EXTERNALSYM NETBIOS_NAME_LEN}

//
// Value to be used with APIs which have a "preferred maximum length"
// parameter.  This value indicates that the API should just allocate
// "as much as it takes."
//

  MAX_PREFERRED_LENGTH = DWORD(-1);
  {$EXTERNALSYM MAX_PREFERRED_LENGTH}

//
//        Constants used with encryption
//

  CRYPT_KEY_LEN      = 7;
  {$EXTERNALSYM CRYPT_KEY_LEN}
  CRYPT_TXT_LEN      = 8;
  {$EXTERNALSYM CRYPT_TXT_LEN}
  ENCRYPTED_PWLEN    = 16;
  {$EXTERNALSYM ENCRYPTED_PWLEN}
  SESSION_PWLEN      = 24;
  {$EXTERNALSYM SESSION_PWLEN}
  SESSION_CRYPT_KLEN = 21;
  {$EXTERNALSYM SESSION_CRYPT_KLEN}

//
//  Value to be used with SetInfo calls to allow setting of all
//  settable parameters (parmnum zero option)
//

  PARMNUM_ALL = 0;
  {$EXTERNALSYM PARMNUM_ALL}

  PARM_ERROR_UNKNOWN     = DWORD(-1);
  {$EXTERNALSYM PARM_ERROR_UNKNOWN}
  PARM_ERROR_NONE        = 0;
  {$EXTERNALSYM PARM_ERROR_NONE}
  PARMNUM_BASE_INFOLEVEL = 1000;
  {$EXTERNALSYM PARMNUM_BASE_INFOLEVEL}

type
  LMSTR = LPWSTR;
  {$EXTERNALSYM LMSTR}
  LMCSTR = LPCWSTR;
  {$EXTERNALSYM LMCSTR}
  PLMSTR = ^LMSTR;
  {$NODEFINE PLMSTR}

//
//        Message File Names
//

const
  MESSAGE_FILENAME  = 'NETMSG';
  {$EXTERNALSYM MESSAGE_FILENAME}
  OS2MSG_FILENAME   = 'BASE';
  {$EXTERNALSYM OS2MSG_FILENAME}
  HELP_MSG_FILENAME = 'NETH';
  {$EXTERNALSYM HELP_MSG_FILENAME}

// ** INTERNAL_ONLY **

// The backup message file named here is a duplicate of net.msg. It
// is not shipped with the product, but is used at buildtime to
// msgbind certain messages to netapi.dll and some of the services.
// This allows for OEMs to modify the message text in net.msg and
// have those changes show up.        Only in case there is an error in
// retrieving the messages from net.msg do we then get the bound
// messages out of bak.msg (really out of the message segment).

  BACKUP_MSG_FILENAME = 'BAK.MSG';
  {$EXTERNALSYM BACKUP_MSG_FILENAME}

// ** END_INTERNAL **

//
// Keywords used in Function Prototypes
//

type
  NET_API_STATUS = DWORD;
  {$EXTERNALSYM NET_API_STATUS}
  TNetApiStatus = NET_API_STATUS;

//
// The platform ID indicates the levels to use for platform-specific
// information.
//

const
  PLATFORM_ID_DOS = 300;
  {$EXTERNALSYM PLATFORM_ID_DOS}
  PLATFORM_ID_OS2 = 400;
  {$EXTERNALSYM PLATFORM_ID_OS2}
  PLATFORM_ID_NT  = 500;
  {$EXTERNALSYM PLATFORM_ID_NT}
  PLATFORM_ID_OSF = 600;
  {$EXTERNALSYM PLATFORM_ID_OSF}
  PLATFORM_ID_VMS = 700;
  {$EXTERNALSYM PLATFORM_ID_VMS}

//
//      There message numbers assigned to different LANMAN components
//      are as defined below.
//
//      lmerr.h:        2100 - 2999     NERR_BASE
//      alertmsg.h:     3000 - 3049     ALERT_BASE
//      lmsvc.h:        3050 - 3099     SERVICE_BASE
//      lmerrlog.h:     3100 - 3299     ERRLOG_BASE
//      msgtext.h:      3300 - 3499     MTXT_BASE
//      apperr.h:       3500 - 3999     APPERR_BASE
//      apperrfs.h:     4000 - 4299     APPERRFS_BASE
//      apperr2.h:      4300 - 5299     APPERR2_BASE
//      ncberr.h:       5300 - 5499     NRCERR_BASE
//      alertmsg.h:     5500 - 5599     ALERT2_BASE
//      lmsvc.h:        5600 - 5699     SERVICE2_BASE
//      lmerrlog.h      5700 - 5899     ERRLOG2_BASE
//

  MIN_LANMAN_MESSAGE_ID = NERR_BASE;
  {$EXTERNALSYM MIN_LANMAN_MESSAGE_ID}
  MAX_LANMAN_MESSAGE_ID = 5899;
  {$EXTERNALSYM MAX_LANMAN_MESSAGE_ID}

// line 59

//
// Function Prototypes - User
//



//
//  Data Structures - User
//

type
  LPUSER_INFO_0 = ^USER_INFO_0;
  {$EXTERNALSYM LPUSER_INFO_0}
  PUSER_INFO_0 = ^USER_INFO_0;
  {$EXTERNALSYM PUSER_INFO_0}
  _USER_INFO_0 = record
    usri0_name: LPWSTR;
  end;
  {$EXTERNALSYM _USER_INFO_0}
  USER_INFO_0 = _USER_INFO_0;
  {$EXTERNALSYM USER_INFO_0}
  TUserInfo0 = USER_INFO_0;
  PUserInfo0 = PUSER_INFO_0;

  LPUSER_INFO_1 = ^USER_INFO_1;
  {$EXTERNALSYM LPUSER_INFO_1}
  PUSER_INFO_1 = ^USER_INFO_1;
  {$EXTERNALSYM PUSER_INFO_1}
  _USER_INFO_1 = record
    usri1_name: LPWSTR;
    usri1_password: LPWSTR;
    usri1_password_age: DWORD;
    usri1_priv: DWORD;
    usri1_home_dir: LPWSTR;
    usri1_comment: LPWSTR;
    usri1_flags: DWORD;
    usri1_script_path: LPWSTR;
  end;
  {$EXTERNALSYM _USER_INFO_1}
  USER_INFO_1 = _USER_INFO_1;
  {$EXTERNALSYM USER_INFO_1}
  TUserInfo1 = USER_INFO_1;
  PUserInfo1 = PUSER_INFO_1;

  LPUSER_INFO_2 = ^USER_INFO_2;
  {$EXTERNALSYM LPUSER_INFO_2}
  PUSER_INFO_2 = ^USER_INFO_2;
  {$EXTERNALSYM PUSER_INFO_2}
  _USER_INFO_2 = record
    usri2_name: LPWSTR;
    usri2_password: LPWSTR;
    usri2_password_age: DWORD;
    usri2_priv: DWORD;
    usri2_home_dir: LPWSTR;
    usri2_comment: LPWSTR;
    usri2_flags: DWORD;
    usri2_script_path: LPWSTR;
    usri2_auth_flags: DWORD;
    usri2_full_name: LPWSTR;
    usri2_usr_comment: LPWSTR;
    usri2_parms: LPWSTR;
    usri2_workstations: LPWSTR;
    usri2_last_logon: DWORD;
    usri2_last_logoff: DWORD;
    usri2_acct_expires: DWORD;
    usri2_max_storage: DWORD;
    usri2_units_per_week: DWORD;
    usri2_logon_hours: PBYTE;
    usri2_bad_pw_count: DWORD;
    usri2_num_logons: DWORD;
    usri2_logon_server: LPWSTR;
    usri2_country_code: DWORD;
    usri2_code_page: DWORD;
  end;
  {$EXTERNALSYM _USER_INFO_2}
  USER_INFO_2 = _USER_INFO_2;
  {$EXTERNALSYM USER_INFO_2}
  TUserInfo2 = USER_INFO_2;
  PUserInfo2 = puser_info_2;

// line 799

//
// Special Values and Constants - User
//

//
//  Bit masks for field usriX_flags of USER_INFO_X (X = 0/1).
//

const
  UF_SCRIPT                          = $0001;
  {$EXTERNALSYM UF_SCRIPT}
  UF_ACCOUNTDISABLE                  = $0002;
  {$EXTERNALSYM UF_ACCOUNTDISABLE}
  UF_HOMEDIR_REQUIRED                = $0008;
  {$EXTERNALSYM UF_HOMEDIR_REQUIRED}
  UF_LOCKOUT                         = $0010;
  {$EXTERNALSYM UF_LOCKOUT}
  UF_PASSWD_NOTREQD                  = $0020;
  {$EXTERNALSYM UF_PASSWD_NOTREQD}
  UF_PASSWD_CANT_CHANGE              = $0040;
  {$EXTERNALSYM UF_PASSWD_CANT_CHANGE}
  UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = $0080;
  {$EXTERNALSYM UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED}

//
// Account type bits as part of usri_flags.
//

  UF_TEMP_DUPLICATE_ACCOUNT    = $0100;
  {$EXTERNALSYM UF_TEMP_DUPLICATE_ACCOUNT}
  UF_NORMAL_ACCOUNT            = $0200;
  {$EXTERNALSYM UF_NORMAL_ACCOUNT}
  UF_INTERDOMAIN_TRUST_ACCOUNT = $0800;
  {$EXTERNALSYM UF_INTERDOMAIN_TRUST_ACCOUNT}
  UF_WORKSTATION_TRUST_ACCOUNT = $1000;
  {$EXTERNALSYM UF_WORKSTATION_TRUST_ACCOUNT}
  UF_SERVER_TRUST_ACCOUNT      = $2000;
  {$EXTERNALSYM UF_SERVER_TRUST_ACCOUNT}

  UF_MACHINE_ACCOUNT_MASK = UF_INTERDOMAIN_TRUST_ACCOUNT or UF_WORKSTATION_TRUST_ACCOUNT or UF_SERVER_TRUST_ACCOUNT;
  {$EXTERNALSYM UF_MACHINE_ACCOUNT_MASK}

  UF_ACCOUNT_TYPE_MASK = UF_TEMP_DUPLICATE_ACCOUNT or UF_NORMAL_ACCOUNT or
    UF_INTERDOMAIN_TRUST_ACCOUNT or UF_WORKSTATION_TRUST_ACCOUNT or UF_SERVER_TRUST_ACCOUNT;
  {$EXTERNALSYM UF_ACCOUNT_TYPE_MASK}

  UF_DONT_EXPIRE_PASSWD                     = $10000;
  {$EXTERNALSYM UF_DONT_EXPIRE_PASSWD}
  UF_MNS_LOGON_ACCOUNT                      = $20000;
  {$EXTERNALSYM UF_MNS_LOGON_ACCOUNT}
  UF_SMARTCARD_REQUIRED                     = $40000;
  {$EXTERNALSYM UF_SMARTCARD_REQUIRED}
  UF_TRUSTED_FOR_DELEGATION                 = $80000;
  {$EXTERNALSYM UF_TRUSTED_FOR_DELEGATION}
  UF_NOT_DELEGATED                          = $100000;
  {$EXTERNALSYM UF_NOT_DELEGATED}
  UF_USE_DES_KEY_ONLY                       = $200000;
  {$EXTERNALSYM UF_USE_DES_KEY_ONLY}
  UF_DONT_REQUIRE_PREAUTH                   = $400000;
  {$EXTERNALSYM UF_DONT_REQUIRE_PREAUTH}
  UF_PASSWORD_EXPIRED                       = DWORD($800000);
  {$EXTERNALSYM UF_PASSWORD_EXPIRED}
  UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = $1000000;
  {$EXTERNALSYM UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION}


  UF_SETTABLE_BITS =
    UF_SCRIPT or
    UF_ACCOUNTDISABLE or
    UF_LOCKOUT or
    UF_HOMEDIR_REQUIRED or
    UF_PASSWD_NOTREQD or
    UF_PASSWD_CANT_CHANGE or
    UF_ACCOUNT_TYPE_MASK or
    UF_DONT_EXPIRE_PASSWD or
    UF_MNS_LOGON_ACCOUNT or
    UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED or
    UF_SMARTCARD_REQUIRED or
    UF_TRUSTED_FOR_DELEGATION or
    UF_NOT_DELEGATED or
    UF_USE_DES_KEY_ONLY or
    UF_DONT_REQUIRE_PREAUTH or
    UF_PASSWORD_EXPIRED or
    UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
  {$EXTERNALSYM UF_SETTABLE_BITS}

// line 1056

//
//  For SetInfo call (parmnum 0) when password change not required
//

  NULL_USERSETINFO_PASSWD = '              ';
  {$EXTERNALSYM NULL_USERSETINFO_PASSWD}

  TIMEQ_FOREVER             = ULONG(-1);
  {$EXTERNALSYM TIMEQ_FOREVER}
  USER_MAXSTORAGE_UNLIMITED = ULONG(-1);
  {$EXTERNALSYM USER_MAXSTORAGE_UNLIMITED}
  USER_NO_LOGOFF            = ULONG(-1);
  {$EXTERNALSYM USER_NO_LOGOFF}
  UNITS_PER_DAY             = 24;
  {$EXTERNALSYM UNITS_PER_DAY}
  UNITS_PER_WEEK            = UNITS_PER_DAY * 7;
  {$EXTERNALSYM UNITS_PER_WEEK}

//
// Privilege levels (USER_INFO_X field usriX_priv (X = 0/1)).
//

  USER_PRIV_MASK  = $3;
  {$EXTERNALSYM USER_PRIV_MASK}
  USER_PRIV_GUEST = 0;
  {$EXTERNALSYM USER_PRIV_GUEST}
  USER_PRIV_USER  = 1;
  {$EXTERNALSYM USER_PRIV_USER}
  USER_PRIV_ADMIN = 2;
  {$EXTERNALSYM USER_PRIV_ADMIN}

// line 1177

//
// Group Class
//

//
// Function Prototypes
//


//
//  Data Structures - Group
//

type
  LPGROUP_INFO_0 = ^GROUP_INFO_0;
  {$EXTERNALSYM LPGROUP_INFO_0}
  PGROUP_INFO_0 = ^GROUP_INFO_0;
  {$EXTERNALSYM PGROUP_INFO_0}
  _GROUP_INFO_0 = record
    grpi0_name: LPWSTR;
  end;
  {$EXTERNALSYM _GROUP_INFO_0}
  GROUP_INFO_0 = _GROUP_INFO_0;
  {$EXTERNALSYM GROUP_INFO_0}
  TGroupInfo0 = GROUP_INFO_0;
  PGroupInfo0 = PGROUP_INFO_0;

  LPGROUP_INFO_1 = ^GROUP_INFO_1;
  {$EXTERNALSYM LPGROUP_INFO_1}
  PGROUP_INFO_1 = ^GROUP_INFO_1;
  {$EXTERNALSYM PGROUP_INFO_1}
  _GROUP_INFO_1 = record
    grpi1_name: LPWSTR;
    grpi1_comment: LPWSTR;
  end;
  {$EXTERNALSYM _GROUP_INFO_1}
  GROUP_INFO_1 = _GROUP_INFO_1;
  {$EXTERNALSYM GROUP_INFO_1}
  TGroupInfo1 = GROUP_INFO_1;
  PGroupInfo1 = PGROUP_INFO_1;

// line 1380

//
// LocalGroup Class
//

//
// Function Prototypes
//



//
//  Data Structures - LocalGroup
//

type
  LPLOCALGROUP_INFO_0 = ^LOCALGROUP_INFO_0;
  {$EXTERNALSYM LPLOCALGROUP_INFO_0}
  PLOCALGROUP_INFO_0 = ^LOCALGROUP_INFO_0;
  {$EXTERNALSYM PLOCALGROUP_INFO_0}
  _LOCALGROUP_INFO_0 = record
    lgrpi0_name: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_INFO_0}
  LOCALGROUP_INFO_0 = _LOCALGROUP_INFO_0;
  {$EXTERNALSYM LOCALGROUP_INFO_0}
  TLocalGroupInfo0 = LOCALGROUP_INFO_0;
  PLocalGroupInfo0 = PLOCALGROUP_INFO_0;

  LPLOCALGROUP_INFO_1 = ^LOCALGROUP_INFO_1;
  {$EXTERNALSYM LPLOCALGROUP_INFO_1}
  PLOCALGROUP_INFO_1 = ^LOCALGROUP_INFO_1;
  {$EXTERNALSYM PLOCALGROUP_INFO_1}
  _LOCALGROUP_INFO_1 = record
    lgrpi1_name: LPWSTR;
    lgrpi1_comment: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_INFO_1}
  LOCALGROUP_INFO_1 = _LOCALGROUP_INFO_1;
  {$EXTERNALSYM LOCALGROUP_INFO_1}
  TLocalGroupInfo1 = LOCALGROUP_INFO_1;
  PLocalGroupInfo1 = PLOCALGROUP_INFO_1;

  LPLOCALGROUP_INFO_1002 = ^LOCALGROUP_INFO_1002;
  {$EXTERNALSYM LPLOCALGROUP_INFO_1002}
  PLOCALGROUP_INFO_1002 = ^LOCALGROUP_INFO_1002;
  {$EXTERNALSYM PLOCALGROUP_INFO_1002}
  _LOCALGROUP_INFO_1002 = record
    lgrpi1002_comment: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_INFO_1002}
  LOCALGROUP_INFO_1002 = _LOCALGROUP_INFO_1002;
  {$EXTERNALSYM LOCALGROUP_INFO_1002}
  TLocalGroupInfo1002 = LOCALGROUP_INFO_1002;
  PLocalGroupInfo1002 = PLOCALGROUP_INFO_1002;

  LPLOCALGROUP_MEMBERS_INFO_0 = ^LOCALGROUP_MEMBERS_INFO_0;
  {$EXTERNALSYM LPLOCALGROUP_MEMBERS_INFO_0}
  PLOCALGROUP_MEMBERS_INFO_0 = ^LOCALGROUP_MEMBERS_INFO_0;
  {$EXTERNALSYM PLOCALGROUP_MEMBERS_INFO_0}
  _LOCALGROUP_MEMBERS_INFO_0 = record
    lgrmi0_sid: PSID;
  end;
  {$EXTERNALSYM _LOCALGROUP_MEMBERS_INFO_0}
  LOCALGROUP_MEMBERS_INFO_0 = _LOCALGROUP_MEMBERS_INFO_0;
  {$EXTERNALSYM LOCALGROUP_MEMBERS_INFO_0}
  TLocalGroupMembersInfo0 = LOCALGROUP_MEMBERS_INFO_0;
  PLocalGroupMembersInfo0 = PLOCALGROUP_MEMBERS_INFO_0;

  LPLOCALGROUP_MEMBERS_INFO_1 = ^LOCALGROUP_MEMBERS_INFO_1;
  {$EXTERNALSYM LPLOCALGROUP_MEMBERS_INFO_1}
  PLOCALGROUP_MEMBERS_INFO_1 = ^LOCALGROUP_MEMBERS_INFO_1;
  {$EXTERNALSYM PLOCALGROUP_MEMBERS_INFO_1}
  _LOCALGROUP_MEMBERS_INFO_1 = record
    lgrmi1_sid: PSID;
    lgrmi1_sidusage: SID_NAME_USE;
    lgrmi1_name: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_MEMBERS_INFO_1}
  LOCALGROUP_MEMBERS_INFO_1 = _LOCALGROUP_MEMBERS_INFO_1;
  {$EXTERNALSYM LOCALGROUP_MEMBERS_INFO_1}
  TLocalGroupMembersInfo1 = LOCALGROUP_MEMBERS_INFO_1;
  PLocalGroupMembersInfo1 = PLOCALGROUP_MEMBERS_INFO_1;

  LPLOCALGROUP_MEMBERS_INFO_2 = ^LOCALGROUP_MEMBERS_INFO_2;
  {$EXTERNALSYM LPLOCALGROUP_MEMBERS_INFO_2}
  PLOCALGROUP_MEMBERS_INFO_2 = ^LOCALGROUP_MEMBERS_INFO_2;
  {$EXTERNALSYM PLOCALGROUP_MEMBERS_INFO_2}
  _LOCALGROUP_MEMBERS_INFO_2 = record
    lgrmi2_sid: PSID;
    lgrmi2_sidusage: SID_NAME_USE;
    lgrmi2_domainandname: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_MEMBERS_INFO_2}
  LOCALGROUP_MEMBERS_INFO_2 = _LOCALGROUP_MEMBERS_INFO_2;
  {$EXTERNALSYM LOCALGROUP_MEMBERS_INFO_2}
  TLocalGroupMembersInfo2 = LOCALGROUP_MEMBERS_INFO_2;
  PLocalGroupMembersInfo2 = PLOCALGROUP_MEMBERS_INFO_2;

  LPLOCALGROUP_MEMBERS_INFO_3 = ^LOCALGROUP_MEMBERS_INFO_3;
  {$EXTERNALSYM LPLOCALGROUP_MEMBERS_INFO_3}
  PLOCALGROUP_MEMBERS_INFO_3 = ^LOCALGROUP_MEMBERS_INFO_3;
  {$EXTERNALSYM PLOCALGROUP_MEMBERS_INFO_3}
  _LOCALGROUP_MEMBERS_INFO_3 = record
    lgrmi3_domainandname: LPWSTR;
  end;
  {$EXTERNALSYM _LOCALGROUP_MEMBERS_INFO_3}
  LOCALGROUP_MEMBERS_INFO_3 = _LOCALGROUP_MEMBERS_INFO_3;
  {$EXTERNALSYM LOCALGROUP_MEMBERS_INFO_3}
  TLocalGroupMembersInfo3 = LOCALGROUP_MEMBERS_INFO_3;
  PLocalGroupMembersInfo3 = PLOCALGROUP_MEMBERS_INFO_3;

type
  _WKSTA_INFO_100 = record
    wki100_platform_id: DWORD;
    wki100_computername: LMSTR;
    wki100_langroup: LMSTR;
    wki100_ver_major: DWORD;
    wki100_ver_minor: DWORD;
  end;
  {$EXTERNALSYM _WKSTA_INFO_100}
  WKSTA_INFO_100 = _WKSTA_INFO_100;
  {$EXTERNALSYM WKSTA_INFO_100}
  PWKSTA_INFO_100 = ^_WKSTA_INFO_100;
  {$EXTERNALSYM PWKSTA_INFO_100}
  LPWKSTA_INFO_100 = ^_WKSTA_INFO_100;
  {$EXTERNALSYM LPWKSTA_INFO_100}

(****************************************************************
 *                                                              *
 *              Data structure templates                        *
 *                                                              *
 ****************************************************************)

const
  NCBNAMSZ = 16;  // absolute length of a net name
  {$EXTERNALSYM NCBNAMSZ}
  MAX_LANA = 254; // lana's in range 0 to MAX_LANA inclusive
  {$EXTERNALSYM MAX_LANA}

//
// Network Control Block
//

type
  PNCB = ^NCB;

  TNcbPost = procedure (P: PNCB); stdcall;

  _NCB = record
    ncb_command: UCHAR;  // command code
    ncb_retcode: UCHAR;  // return code
    ncb_lsn: UCHAR;      // local session number
    ncb_num: UCHAR;      // number of our network name
    ncb_buffer: PUCHAR;  // address of message buffer
    ncb_length: Word;    // size of message buffer
    ncb_callname: array [0..NCBNAMSZ - 1] of UCHAR; // blank-padded name of remote
    ncb_name: array [0..NCBNAMSZ - 1] of UCHAR;     // our blank-padded netname
    ncb_rto: UCHAR;      // rcv timeout/retry count
    ncb_sto: UCHAR;      // send timeout/sys timeout
    ncb_post: TNcbPost;  // POST routine address
    ncb_lana_num: UCHAR; // lana (adapter) number
    ncb_cmd_cplt: UCHAR; // 0xff => commmand pending
    {$IFDEF WIN64}
    ncb_reserve: array [0..17] of UCHAR; // reserved, used by BIOS
    {$ELSE ~WIN64}
    ncb_reserve: array [0..9] of UCHAR;  // reserved, used by BIOS
    {$ENDIF ~WIN64}
    ncb_event: THandle;   // HANDLE to Win32 event which
                         // will be set to the signalled
                         // state when an ASYNCH command
                         // completes
  end;
  {$EXTERNALSYM _NCB}
  NCB = _NCB;
  {$EXTERNALSYM NCB}
  TNcb = NCB;

//
//  Structure returned to the NCB command NCBASTAT is ADAPTER_STATUS followed
//  by an array of NAME_BUFFER structures.
//
type
  _ADAPTER_STATUS = record
    adapter_address: array [0..5] of UCHAR;
    rev_major: UCHAR;
    reserved0: UCHAR;
    adapter_type: UCHAR;
    rev_minor: UCHAR;
    duration: WORD;
    frmr_recv: WORD;
    frmr_xmit: WORD;
    iframe_recv_err: WORD;
    xmit_aborts: WORD;
    xmit_success: DWORD;
    recv_success: DWORD;
    iframe_xmit_err: WORD;
    recv_buff_unavail: WORD;
    t1_timeouts: WORD;
    ti_timeouts: WORD;
    reserved1: DWORD;
    free_ncbs: WORD;
    max_cfg_ncbs: WORD;
    max_ncbs: WORD;
    xmit_buf_unavail: WORD;
    max_dgram_size: WORD;
    pending_sess: WORD;
    max_cfg_sess: WORD;
    max_sess: WORD;
    max_sess_pkt_size: WORD;
    name_count: WORD;
  end;
  {$EXTERNALSYM _ADAPTER_STATUS}
  ADAPTER_STATUS = _ADAPTER_STATUS;
  {$EXTERNALSYM ADAPTER_STATUS}
  PADAPTER_STATUS = ^ADAPTER_STATUS;
  {$EXTERNALSYM PADAPTER_STATUS}
  TAdapterStatus = ADAPTER_STATUS;
  PAdapterStatus = PADAPTER_STATUS;

  _NAME_BUFFER = record
    name: array [0..NCBNAMSZ - 1] of UCHAR;
    name_num: UCHAR;
    name_flags: UCHAR;
  end;
  {$EXTERNALSYM _NAME_BUFFER}
  NAME_BUFFER = _NAME_BUFFER;
  {$EXTERNALSYM NAME_BUFFER}
  PNAME_BUFFER = ^NAME_BUFFER;
  {$EXTERNALSYM PNAME_BUFFER}
  TNameBuffer = NAME_BUFFER;
  PNameBuffer = PNAME_BUFFER;

//  values for name_flags bits.

const
  NAME_FLAGS_MASK = $87;
  {$EXTERNALSYM NAME_FLAGS_MASK}

  GROUP_NAME  = $80;
  {$EXTERNALSYM GROUP_NAME}
  UNIQUE_NAME = $00;
  {$EXTERNALSYM UNIQUE_NAME}

  REGISTERING     = $00;
  {$EXTERNALSYM REGISTERING}
  REGISTERED      = $04;
  {$EXTERNALSYM REGISTERED}
  DEREGISTERED    = $05;
  {$EXTERNALSYM DEREGISTERED}
  DUPLICATE       = $06;
  {$EXTERNALSYM DUPLICATE}
  DUPLICATE_DEREG = $07;
  {$EXTERNALSYM DUPLICATE_DEREG}

//
//  Structure returned to the NCB command NCBSSTAT is SESSION_HEADER followed
//  by an array of SESSION_BUFFER structures. If the NCB_NAME starts with an
//  asterisk then an array of these structures is returned containing the
//  status for all names.
//

type
  _SESSION_HEADER = record
    sess_name: UCHAR;
    num_sess: UCHAR;
    rcv_dg_outstanding: UCHAR;
    rcv_any_outstanding: UCHAR;
  end;
  {$EXTERNALSYM _SESSION_HEADER}
  SESSION_HEADER = _SESSION_HEADER;
  {$EXTERNALSYM SESSION_HEADER}
  PSESSION_HEADER = ^SESSION_HEADER;
  {$EXTERNALSYM PSESSION_HEADER}
  TSessionHeader = SESSION_HEADER;
  PSessionHeader = PSESSION_HEADER;

  _SESSION_BUFFER = record
    lsn: UCHAR;
    state: UCHAR;
    local_name: array [0..NCBNAMSZ - 1] of UCHAR;
    remote_name: array [0..NCBNAMSZ - 1] of UCHAR;
    rcvs_outstanding: UCHAR;
    sends_outstanding: UCHAR;
  end;
  {$EXTERNALSYM _SESSION_BUFFER}
  SESSION_BUFFER = _SESSION_BUFFER;
  {$EXTERNALSYM SESSION_BUFFER}
  PSESSION_BUFFER = ^SESSION_BUFFER;
  {$EXTERNALSYM PSESSION_BUFFER}
  TSessionBuffer = SESSION_BUFFER;
  PSessionBuffer = PSESSION_BUFFER;

//  Values for state

const
  LISTEN_OUTSTANDING  = $01;
  {$EXTERNALSYM LISTEN_OUTSTANDING}
  CALL_PENDING        = $02;
  {$EXTERNALSYM CALL_PENDING}
  SESSION_ESTABLISHED = $03;
  {$EXTERNALSYM SESSION_ESTABLISHED}
  HANGUP_PENDING      = $04;
  {$EXTERNALSYM HANGUP_PENDING}
  HANGUP_COMPLETE     = $05;
  {$EXTERNALSYM HANGUP_COMPLETE}
  SESSION_ABORTED     = $06;
  {$EXTERNALSYM SESSION_ABORTED}

//
//  Structure returned to the NCB command NCBENUM.
//
//  On a system containing lana's 0, 2 and 3, a structure with
//  length =3, lana[0]=0, lana[1]=2 and lana[2]=3 will be returned.
//

type
  _LANA_ENUM = record
    length: UCHAR; // Number of valid entries in lana[]
    lana: array [0..MAX_LANA] of UCHAR;
  end;
  {$EXTERNALSYM _LANA_ENUM}
  LANA_ENUM = _LANA_ENUM;
  {$EXTERNALSYM LANA_ENUM}
  PLANA_ENUM = ^LANA_ENUM;
  {$EXTERNALSYM PLANA_ENUM}
  TLanaEnum = LANA_ENUM;
  PLanaEnum = PLANA_ENUM;

//
//  Structure returned to the NCB command NCBFINDNAME is FIND_NAME_HEADER followed
//  by an array of FIND_NAME_BUFFER structures.
//

type
  _FIND_NAME_HEADER = record
    node_count: WORD;
    reserved: UCHAR;
    unique_group: UCHAR;
  end;
  {$EXTERNALSYM _FIND_NAME_HEADER}
  FIND_NAME_HEADER = _FIND_NAME_HEADER;
  {$EXTERNALSYM FIND_NAME_HEADER}
  PFIND_NAME_HEADER = ^FIND_NAME_HEADER;
  {$EXTERNALSYM PFIND_NAME_HEADER}
  TFindNameHeader = FIND_NAME_HEADER;
  PFindNameHeader = PFIND_NAME_HEADER;

  _FIND_NAME_BUFFER = record
    length: UCHAR;
    access_control: UCHAR;
    frame_control: UCHAR;
    destination_addr: array [0..5] of UCHAR;
    source_addr: array [0..5] of UCHAR;
    routing_info: array [0..17] of UCHAR;
  end;
  {$EXTERNALSYM _FIND_NAME_BUFFER}
  FIND_NAME_BUFFER = _FIND_NAME_BUFFER;
  {$EXTERNALSYM FIND_NAME_BUFFER}
  PFIND_NAME_BUFFER = ^FIND_NAME_BUFFER;
  {$EXTERNALSYM PFIND_NAME_BUFFER}
  TFindNameBuffer = FIND_NAME_BUFFER;
  PFindNameBuffer = PFIND_NAME_BUFFER;

//
//  Structure provided with NCBACTION. The purpose of NCBACTION is to provide
//  transport specific extensions to netbios.
//

  _ACTION_HEADER = record
    transport_id: ULONG;
    action_code: USHORT;
    reserved: USHORT;
  end;
  {$EXTERNALSYM _ACTION_HEADER}
  ACTION_HEADER = _ACTION_HEADER;
  {$EXTERNALSYM ACTION_HEADER}
  PACTION_HEADER = ^ACTION_HEADER;
  {$EXTERNALSYM PACTION_HEADER}
  TActionHeader = ACTION_HEADER;
  PActionHeader = PACTION_HEADER;

//  Values for transport_id

const
  ALL_TRANSPORTS = 'M'#0#0#0;
  {$EXTERNALSYM ALL_TRANSPORTS}
  MS_NBF         = 'MNBF';
  {$EXTERNALSYM MS_NBF}

(****************************************************************
 *                                                              *
 *              Special values and constants                    *
 *                                                              *
 ****************************************************************)

//
//      NCB Command codes
//

const
  NCBCALL        = $10; // NCB CALL
  {$EXTERNALSYM NCBCALL}
  NCBLISTEN      = $11; // NCB LISTEN
  {$EXTERNALSYM NCBLISTEN}
  NCBHANGUP      = $12; // NCB HANG UP
  {$EXTERNALSYM NCBHANGUP}
  NCBSEND        = $14; // NCB SEND
  {$EXTERNALSYM NCBSEND}
  NCBRECV        = $15; // NCB RECEIVE
  {$EXTERNALSYM NCBRECV}
  NCBRECVANY     = $16; // NCB RECEIVE ANY
  {$EXTERNALSYM NCBRECVANY}
  NCBCHAINSEND   = $17; // NCB CHAIN SEND
  {$EXTERNALSYM NCBCHAINSEND}
  NCBDGSEND      = $20; // NCB SEND DATAGRAM
  {$EXTERNALSYM NCBDGSEND}
  NCBDGRECV      = $21; // NCB RECEIVE DATAGRAM
  {$EXTERNALSYM NCBDGRECV}
  NCBDGSENDBC    = $22; // NCB SEND BROADCAST DATAGRAM
  {$EXTERNALSYM NCBDGSENDBC}
  NCBDGRECVBC    = $23; // NCB RECEIVE BROADCAST DATAGRAM
  {$EXTERNALSYM NCBDGRECVBC}
  NCBADDNAME     = $30; // NCB ADD NAME
  {$EXTERNALSYM NCBADDNAME}
  NCBDELNAME     = $31; // NCB DELETE NAME
  {$EXTERNALSYM NCBDELNAME}
  NCBRESET       = $32; // NCB RESET
  {$EXTERNALSYM NCBRESET}
  NCBASTAT       = $33; // NCB ADAPTER STATUS
  {$EXTERNALSYM NCBASTAT}
  NCBSSTAT       = $34; // NCB SESSION STATUS
  {$EXTERNALSYM NCBSSTAT}
  NCBCANCEL      = $35; // NCB CANCEL
  {$EXTERNALSYM NCBCANCEL}
  NCBADDGRNAME   = $36; // NCB ADD GROUP NAME
  {$EXTERNALSYM NCBADDGRNAME}
  NCBENUM        = $37; // NCB ENUMERATE LANA NUMBERS
  {$EXTERNALSYM NCBENUM}
  NCBUNLINK      = $70; // NCB UNLINK
  {$EXTERNALSYM NCBUNLINK}
  NCBSENDNA      = $71; // NCB SEND NO ACK
  {$EXTERNALSYM NCBSENDNA}
  NCBCHAINSENDNA = $72; // NCB CHAIN SEND NO ACK
  {$EXTERNALSYM NCBCHAINSENDNA}
  NCBLANSTALERT  = $73; // NCB LAN STATUS ALERT
  {$EXTERNALSYM NCBLANSTALERT}
  NCBACTION      = $77; // NCB ACTION
  {$EXTERNALSYM NCBACTION}
  NCBFINDNAME    = $78; // NCB FIND NAME
  {$EXTERNALSYM NCBFINDNAME}
  NCBTRACE       = $79; // NCB TRACE
  {$EXTERNALSYM NCBTRACE}

  ASYNCH = $80; // high bit set == asynchronous
  {$EXTERNALSYM ASYNCH}

//
//      NCB Return codes
//

  NRC_GOODRET = $00; // good return also returned when ASYNCH request accepted
  {$EXTERNALSYM NRC_GOODRET}
  NRC_BUFLEN      = $01; // illegal buffer length
  {$EXTERNALSYM NRC_BUFLEN}
  NRC_ILLCMD      = $03; // illegal command
  {$EXTERNALSYM NRC_ILLCMD}
  NRC_CMDTMO      = $05; // command timed out
  {$EXTERNALSYM NRC_CMDTMO}
  NRC_INCOMP      = $06; // message incomplete, issue another command
  {$EXTERNALSYM NRC_INCOMP}
  NRC_BADDR       = $07; // illegal buffer address
  {$EXTERNALSYM NRC_BADDR}
  NRC_SNUMOUT     = $08; // session number out of range
  {$EXTERNALSYM NRC_SNUMOUT}
  NRC_NORES       = $09; // no resource available
  {$EXTERNALSYM NRC_NORES}
  NRC_SCLOSED     = $0a; // session closed
  {$EXTERNALSYM NRC_SCLOSED}
  NRC_CMDCAN      = $0b; // command cancelled
  {$EXTERNALSYM NRC_CMDCAN}
  NRC_DUPNAME     = $0d; // duplicate name
  {$EXTERNALSYM NRC_DUPNAME}
  NRC_NAMTFUL     = $0e; // name table full
  {$EXTERNALSYM NRC_NAMTFUL}
  NRC_ACTSES      = $0f; // no deletions, name has active sessions
  {$EXTERNALSYM NRC_ACTSES}
  NRC_LOCTFUL     = $11; // local session table full
  {$EXTERNALSYM NRC_LOCTFUL}
  NRC_REMTFUL     = $12; // remote session table full
  {$EXTERNALSYM NRC_REMTFUL}
  NRC_ILLNN       = $13; // illegal name number
  {$EXTERNALSYM NRC_ILLNN}
  NRC_NOCALL      = $14; // no callname
  {$EXTERNALSYM NRC_NOCALL}
  NRC_NOWILD      = $15; // cannot put * in NCB_NAME
  {$EXTERNALSYM NRC_NOWILD}
  NRC_INUSE       = $16; // name in use on remote adapter
  {$EXTERNALSYM NRC_INUSE}
  NRC_NAMERR      = $17; // name deleted
  {$EXTERNALSYM NRC_NAMERR}
  NRC_SABORT      = $18; // session ended abnormally
  {$EXTERNALSYM NRC_SABORT}
  NRC_NAMCONF     = $19; // name conflict detected
  {$EXTERNALSYM NRC_NAMCONF}
  NRC_IFBUSY      = $21; // interface busy, IRET before retrying
  {$EXTERNALSYM NRC_IFBUSY}
  NRC_TOOMANY     = $22; // too many commands outstanding, retry later
  {$EXTERNALSYM NRC_TOOMANY}
  NRC_BRIDGE      = $23; // ncb_lana_num field invalid
  {$EXTERNALSYM NRC_BRIDGE}
  NRC_CANOCCR     = $24; // command completed while cancel occurring
  {$EXTERNALSYM NRC_CANOCCR}
  NRC_CANCEL      = $26; // command not valid to cancel
  {$EXTERNALSYM NRC_CANCEL}
  NRC_DUPENV      = $30; // name defined by anther local process
  {$EXTERNALSYM NRC_DUPENV}
  NRC_ENVNOTDEF   = $34; // environment undefined. RESET required
  {$EXTERNALSYM NRC_ENVNOTDEF}
  NRC_OSRESNOTAV  = $35; // required OS resources exhausted
  {$EXTERNALSYM NRC_OSRESNOTAV}
  NRC_MAXAPPS     = $36; // max number of applications exceeded
  {$EXTERNALSYM NRC_MAXAPPS}
  NRC_NOSAPS      = $37; // no saps available for netbios
  {$EXTERNALSYM NRC_NOSAPS}
  NRC_NORESOURCES = $38; // requested resources are not available
  {$EXTERNALSYM NRC_NORESOURCES}
  NRC_INVADDRESS  = $39; // invalid ncb address or length > segment
  {$EXTERNALSYM NRC_INVADDRESS}
  NRC_INVDDID     = $3B; // invalid NCB DDID
  {$EXTERNALSYM NRC_INVDDID}
  NRC_LOCKFAIL    = $3C; // lock of user area failed
  {$EXTERNALSYM NRC_LOCKFAIL}
  NRC_OPENERR     = $3f; // NETBIOS not loaded
  {$EXTERNALSYM NRC_OPENERR}
  NRC_SYSTEM      = $40; // system error
  {$EXTERNALSYM NRC_SYSTEM}

  NRC_PENDING = $ff; // asynchronous command is not yet finished
  {$EXTERNALSYM NRC_PENDING}

(****************************************************************
 *                                                              *
 *              main user entry point for NetBIOS 3.0           *
 *                                                              *
 * Usage: result = Netbios( pncb );                             *
 ****************************************************************)

type
  PRasDialDlg = ^TRasDialDlg;
  tagRASDIALDLG = packed record
    dwSize: DWORD;
    hwndOwner: HWND;
    dwFlags: DWORD;
    xDlg: Longint;
    yDlg: Longint;
    dwSubEntry: DWORD;
    dwError: DWORD;
    reserved: Longword;
    reserved2: Longword;
  end;
  {$EXTERNALSYM tagRASDIALDLG}
  RASDIALDLG = tagRASDIALDLG;
  {$EXTERNALSYM RASDIALDLG}
  TRasDialDlg = tagRASDIALDLG;


// Reason flags

// Flags used by the various UIs.

const
  SHTDN_REASON_FLAG_COMMENT_REQUIRED          = $01000000;
  {$EXTERNALSYM SHTDN_REASON_FLAG_COMMENT_REQUIRED}
  SHTDN_REASON_FLAG_DIRTY_PROBLEM_ID_REQUIRED = $02000000;
  {$EXTERNALSYM SHTDN_REASON_FLAG_DIRTY_PROBLEM_ID_REQUIRED}
  SHTDN_REASON_FLAG_CLEAN_UI                  = $04000000;
  {$EXTERNALSYM SHTDN_REASON_FLAG_CLEAN_UI}
  SHTDN_REASON_FLAG_DIRTY_UI                  = $08000000;
  {$EXTERNALSYM SHTDN_REASON_FLAG_DIRTY_UI}

// Flags that end up in the event log code.

  SHTDN_REASON_FLAG_USER_DEFINED = $40000000;
  {$EXTERNALSYM SHTDN_REASON_FLAG_USER_DEFINED}
  SHTDN_REASON_FLAG_PLANNED      = DWORD($80000000);
  {$EXTERNALSYM SHTDN_REASON_FLAG_PLANNED}

// Microsoft major reasons.

  SHTDN_REASON_MAJOR_OTHER           = $00000000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_OTHER}
  SHTDN_REASON_MAJOR_NONE            = $00000000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_NONE}
  SHTDN_REASON_MAJOR_HARDWARE        = $00010000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_HARDWARE}
  SHTDN_REASON_MAJOR_OPERATINGSYSTEM = $00020000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_OPERATINGSYSTEM}
  SHTDN_REASON_MAJOR_SOFTWARE        = $00030000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_SOFTWARE}
  SHTDN_REASON_MAJOR_APPLICATION     = $00040000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_APPLICATION}
  SHTDN_REASON_MAJOR_SYSTEM          = $00050000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_SYSTEM}
  SHTDN_REASON_MAJOR_POWER           = $00060000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_POWER}
  SHTDN_REASON_MAJOR_LEGACY_API      = $00070000;
  {$EXTERNALSYM SHTDN_REASON_MAJOR_LEGACY_API}

// Microsoft minor reasons.

  SHTDN_REASON_MINOR_OTHER           = $00000000;
  {$EXTERNALSYM SHTDN_REASON_MINOR_OTHER}
  SHTDN_REASON_MINOR_NONE            = $000000ff;
  {$EXTERNALSYM SHTDN_REASON_MINOR_NONE}
  SHTDN_REASON_MINOR_MAINTENANCE     = $00000001;
  {$EXTERNALSYM SHTDN_REASON_MINOR_MAINTENANCE}
  SHTDN_REASON_MINOR_INSTALLATION    = $00000002;
  {$EXTERNALSYM SHTDN_REASON_MINOR_INSTALLATION}
  SHTDN_REASON_MINOR_UPGRADE         = $00000003;
  {$EXTERNALSYM SHTDN_REASON_MINOR_UPGRADE}
  SHTDN_REASON_MINOR_RECONFIG        = $00000004;
  {$EXTERNALSYM SHTDN_REASON_MINOR_RECONFIG}
  SHTDN_REASON_MINOR_HUNG            = $00000005;
  {$EXTERNALSYM SHTDN_REASON_MINOR_HUNG}
  SHTDN_REASON_MINOR_UNSTABLE        = $00000006;
  {$EXTERNALSYM SHTDN_REASON_MINOR_UNSTABLE}
  SHTDN_REASON_MINOR_DISK            = $00000007;
  {$EXTERNALSYM SHTDN_REASON_MINOR_DISK}
  SHTDN_REASON_MINOR_PROCESSOR       = $00000008;
  {$EXTERNALSYM SHTDN_REASON_MINOR_PROCESSOR}
  SHTDN_REASON_MINOR_NETWORKCARD     = $00000009;
  {$EXTERNALSYM SHTDN_REASON_MINOR_NETWORKCARD}
  SHTDN_REASON_MINOR_POWER_SUPPLY    = $0000000a;
  {$EXTERNALSYM SHTDN_REASON_MINOR_POWER_SUPPLY}
  SHTDN_REASON_MINOR_CORDUNPLUGGED   = $0000000b;
  {$EXTERNALSYM SHTDN_REASON_MINOR_CORDUNPLUGGED}
  SHTDN_REASON_MINOR_ENVIRONMENT     = $0000000c;
  {$EXTERNALSYM SHTDN_REASON_MINOR_ENVIRONMENT}
  SHTDN_REASON_MINOR_HARDWARE_DRIVER = $0000000d;
  {$EXTERNALSYM SHTDN_REASON_MINOR_HARDWARE_DRIVER}
  SHTDN_REASON_MINOR_OTHERDRIVER     = $0000000e;
  {$EXTERNALSYM SHTDN_REASON_MINOR_OTHERDRIVER}
  SHTDN_REASON_MINOR_BLUESCREEN      = $0000000F;
  {$EXTERNALSYM SHTDN_REASON_MINOR_BLUESCREEN}
  SHTDN_REASON_MINOR_SERVICEPACK           = $00000010;
  {$EXTERNALSYM SHTDN_REASON_MINOR_SERVICEPACK}
  SHTDN_REASON_MINOR_HOTFIX                = $00000011;
  {$EXTERNALSYM SHTDN_REASON_MINOR_HOTFIX}
  SHTDN_REASON_MINOR_SECURITYFIX           = $00000012;
  {$EXTERNALSYM SHTDN_REASON_MINOR_SECURITYFIX}
  SHTDN_REASON_MINOR_SECURITY              = $00000013;
  {$EXTERNALSYM SHTDN_REASON_MINOR_SECURITY}
  SHTDN_REASON_MINOR_NETWORK_CONNECTIVITY  = $00000014;
  {$EXTERNALSYM SHTDN_REASON_MINOR_NETWORK_CONNECTIVITY}
  SHTDN_REASON_MINOR_WMI                   = $00000015;
  {$EXTERNALSYM SHTDN_REASON_MINOR_WMI}
  SHTDN_REASON_MINOR_SERVICEPACK_UNINSTALL = $00000016;
  {$EXTERNALSYM SHTDN_REASON_MINOR_SERVICEPACK_UNINSTALL}
  SHTDN_REASON_MINOR_HOTFIX_UNINSTALL      = $00000017;
  {$EXTERNALSYM SHTDN_REASON_MINOR_HOTFIX_UNINSTALL}
  SHTDN_REASON_MINOR_SECURITYFIX_UNINSTALL = $00000018;
  {$EXTERNALSYM SHTDN_REASON_MINOR_SECURITYFIX_UNINSTALL}
  SHTDN_REASON_MINOR_MMC                   = $00000019;
  {$EXTERNALSYM SHTDN_REASON_MINOR_MMC}
  SHTDN_REASON_MINOR_TERMSRV               = $00000020;
  {$EXTERNALSYM SHTDN_REASON_MINOR_TERMSRV}
  SHTDN_REASON_MINOR_DC_PROMOTION          = $00000021;
  {$EXTERNALSYM SHTDN_REASON_MINOR_DC_PROMOTION}
  SHTDN_REASON_MINOR_DC_DEMOTION           = $00000022;
  {$EXTERNALSYM SHTDN_REASON_MINOR_DC_DEMOTION}

  SHTDN_REASON_UNKNOWN = SHTDN_REASON_MINOR_NONE;
  {$EXTERNALSYM SHTDN_REASON_UNKNOWN}
  SHTDN_REASON_LEGACY_API = (SHTDN_REASON_MAJOR_LEGACY_API or SHTDN_REASON_FLAG_PLANNED);
  {$EXTERNALSYM SHTDN_REASON_LEGACY_API}

// This mask cuts out UI flags.

  SHTDN_REASON_VALID_BIT_MASK = DWORD($c0ffffff);
  {$EXTERNALSYM SHTDN_REASON_VALID_BIT_MASK}

// Convenience flags.

  PCLEANUI = (SHTDN_REASON_FLAG_PLANNED or SHTDN_REASON_FLAG_CLEAN_UI);
  {$EXTERNALSYM PCLEANUI}
  UCLEANUI = (SHTDN_REASON_FLAG_CLEAN_UI);
  {$EXTERNALSYM UCLEANUI}
  PDIRTYUI = (SHTDN_REASON_FLAG_PLANNED or SHTDN_REASON_FLAG_DIRTY_UI);
  {$EXTERNALSYM PDIRTYUI}
  UDIRTYUI = (SHTDN_REASON_FLAG_DIRTY_UI);
  {$EXTERNALSYM UDIRTYUI}

const
  CSIDL_LOCAL_APPDATA        = $001C; { <user name>\Local Settings\Application Data (non roaming) }
  CSIDL_COMMON_APPDATA       = $0023; { All Users\Application Data }
  CSIDL_WINDOWS              = $0024; { GetWindowsDirectory() }
  CSIDL_SYSTEM               = $0025; { GetSystemDirectory() }
  CSIDL_PROGRAM_FILES        = $0026; { C:\Program Files }
  CSIDL_MYPICTURES           = $0027; { C:\Program Files\My Pictures }
  CSIDL_PROFILE              = $0028; { USERPROFILE }
  CSIDL_PROGRAM_FILESX86     = $002A; { C:\Program Files (x86)\My Pictures }
  CSIDL_PROGRAM_FILES_COMMON = $002B; { C:\Program Files\Common }
  CSIDL_COMMON_TEMPLATES     = $002D; { All Users\Templates }
  CSIDL_COMMON_DOCUMENTS     = $002E; { All Users\Documents }
  CSIDL_COMMON_ADMINTOOLS    = $002F; { All Users\Start Menu\Programs\Administrative Tools }
  CSIDL_ADMINTOOLS           = $0030; { <user name>\Start Menu\Programs\Administrative Tools }
  CSIDL_CONNECTIONS          = $0031; { Network and Dial-up Connections }
  CSIDL_COMMON_MUSIC         = $0035; { All Users\My Music }
  CSIDL_COMMON_PICTURES      = $0036; { All Users\My Pictures }
  CSIDL_COMMON_VIDEO         = $0037; { All Users\My Video }
  CSIDL_RESOURCES            = $0038; { Resource Direcotry }
  CSIDL_RESOURCES_LOCALIZED  = $0039; { Localized Resource Direcotry }
  CSIDL_COMMON_OEM_LINKS     = $003A; { Links to All Users OEM specific apps }
  CSIDL_CDBURN_AREA          = $003B; { USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning }
  CSIDL_COMPUTERSNEARME      = $003D; { Computers Near Me (computered from Workgroup membership) }

  {$EXTERNALSYM CSIDL_LOCAL_APPDATA}
  {$EXTERNALSYM CSIDL_COMMON_APPDATA}
  {$EXTERNALSYM CSIDL_WINDOWS}
  {$EXTERNALSYM CSIDL_SYSTEM}
  {$EXTERNALSYM CSIDL_PROGRAM_FILES}
  {$EXTERNALSYM CSIDL_MYPICTURES}
  {$EXTERNALSYM CSIDL_PROFILE}
  {$EXTERNALSYM CSIDL_PROGRAM_FILESX86}
  {$EXTERNALSYM CSIDL_PROGRAM_FILES_COMMON}
  {$EXTERNALSYM CSIDL_COMMON_TEMPLATES}
  {$EXTERNALSYM CSIDL_COMMON_DOCUMENTS}
  {$EXTERNALSYM CSIDL_COMMON_ADMINTOOLS}
  {$EXTERNALSYM CSIDL_ADMINTOOLS}
  {$EXTERNALSYM CSIDL_CONNECTIONS}
  {$EXTERNALSYM CSIDL_COMMON_MUSIC}
  {$EXTERNALSYM CSIDL_COMMON_PICTURES}
  {$EXTERNALSYM CSIDL_COMMON_VIDEO}
  {$EXTERNALSYM CSIDL_RESOURCES}
  {$EXTERNALSYM CSIDL_RESOURCES_LOCALIZED}
  {$EXTERNALSYM CSIDL_COMMON_OEM_LINKS}
  {$EXTERNALSYM CSIDL_CDBURN_AREA}
  {$EXTERNALSYM CSIDL_COMPUTERSNEARME}

type
  ITaskbarList = interface(IUnknown)
    ['{56FDF342-FD6D-11D0-958A-006097C9A090}']
    function HrInit: HRESULT; stdcall;
    function AddTab(hwnd: HWND): HRESULT; stdcall;
    function DeleteTab(hwnd: HWND): HRESULT; stdcall;
    function ActivateTab(hwnd: HWND): HRESULT; stdcall;
    function SetActiveAlt(hwnd: HWND): HRESULT; stdcall;
  end;
  {$EXTERNALSYM ITaskbarList}

  ITaskbarList2 = interface(ITaskbarList)
    ['{602D4995-B13A-429B-A66E-1935E44F4317}']
    function MarkFullscreenWindow(hwnd: HWND; fFullscreen: BOOL): HRESULT; stdcall;
  end;
  {$EXTERNALSYM ITaskbarList2}

type
  THUMBBUTTON = record
    dwMask: DWORD;
    iId: UINT;
    iBitmap: UINT;
    hIcon: HICON;
    szTip: packed array[0..259] of WCHAR;
    dwFlags: DWORD;
  end;
  {$EXTERNALSYM THUMBBUTTON}
  tagTHUMBBUTTON = THUMBBUTTON;
  {$EXTERNALSYM tagTHUMBBUTTON}
  TThumbButton = THUMBBUTTON;
  {$EXTERNALSYM TThumbButton}
  PThumbButton = ^TThumbButton;
  {$EXTERNALSYM PThumbButton}

// for ThumbButtons.dwFlags
const
  THBF_ENABLED        = $0000;
  {$EXTERNALSYM THBF_ENABLED}
  THBF_DISABLED       = $0001;
  {$EXTERNALSYM THBF_DISABLED}
  THBF_DISMISSONCLICK = $0002;
  {$EXTERNALSYM THBF_DISMISSONCLICK}
  THBF_NOBACKGROUND   = $0004;
  {$EXTERNALSYM THBF_NOBACKGROUND}
  THBF_HIDDEN         = $0008;
  {$EXTERNALSYM THBF_HIDDEN}
  THBF_NONINTERACTIVE = $0010;
  {$EXTERNALSYM THBF_NONINTERACTIVE}

// for ThumbButton.dwMask
const
  THB_BITMAP          = $0001;
  {$EXTERNALSYM THB_BITMAP}
  THB_ICON            = $0002;
  {$EXTERNALSYM THB_ICON}
  THB_TOOLTIP         = $0004;
  {$EXTERNALSYM THB_TOOLTIP}
  THB_FLAGS           = $0008;
  {$EXTERNALSYM THB_FLAGS}

// wParam for WM_COMMAND message (lParam = Button ID)
const
  THBN_CLICKED        = $1800;
  {$EXTERNALSYM THBN_CLICKED}

// for ITaskBarList3.SetProgressState
const
  TBPF_NOPROGRESS     = 0;
  {$EXTERNALSYM TBPF_NOPROGRESS}
  TBPF_INDETERMINATE  = $1;
  {$EXTERNALSYM TBPF_INDETERMINATE}
  TBPF_NORMAL         = $2;
  {$EXTERNALSYM TBPF_NORMAL}
  TBPF_ERROR          = $4;
  {$EXTERNALSYM TBPF_ERROR}
  TBPF_PAUSED         = $8;
  {$EXTERNALSYM TBPF_PAUSED}

type
  ITaskbarList3 = interface(ITaskbarList2)
    ['{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}']
    function SetProgressValue(hwnd: HWND; ullCompleted: ULONGLONG;
      ullTotal: ULONGLONG): HRESULT; stdcall;
    function SetProgressState(hwnd: HWND; tbpFlags: Integer): HRESULT; stdcall;
    function RegisterTab(hwndTab: HWND; hwndMDI: HWND): HRESULT; stdcall;
    function UnregisterTab(hwndTab: HWND): HRESULT; stdcall;
    function SetTabOrder(hwndTab: HWND; hwndInsertBefore: HWND): HRESULT; stdcall;
    function SetTabActive(hwndTab: HWND; hwndMDI: HWND;
      tbatFlags: Integer): HRESULT; stdcall;
    function ThumbBarAddButtons(hwnd: HWND; cButtons: UINT;
      pButton: PThumbButton): HRESULT; stdcall;
    function ThumbBarUpdateButtons(hwnd: HWND; cButtons: UINT;
      pButton: PThumbButton): HRESULT; stdcall;
    function ThumbBarSetImageList(hwnd: HWND; himl: THandle): HRESULT; stdcall;
    function SetOverlayIcon(hwnd: HWND; hIcon: HICON;
      pszDescription: LPCWSTR): HRESULT; stdcall;
    function SetThumbnailTooltip(hwnd: HWND; pszTip: LPCWSTR): HRESULT; stdcall;
    function SetThumbnailClip(hwnd: HWND; var prcClip: TRect): HRESULT; stdcall;
  end;
  {$EXTERNALSYM ITaskbarList3}

type
  STPFLAG = Integer;
  {$EXTERNALSYM STPFLAG}
const
  STPF_NONE                      = 0;
  {$EXTERNALSYM STPF_NONE}
  STPF_USEAPPTHUMBNAILALWAYS     = $1;
  {$EXTERNALSYM STPF_USEAPPTHUMBNAILALWAYS}
  STPF_USEAPPTHUMBNAILWHENACTIVE = $2;
  {$EXTERNALSYM STPF_USEAPPTHUMBNAILWHENACTIVE}
  STPF_USEAPPPEEKALWAYS          = $4;
  {$EXTERNALSYM STPF_USEAPPPEEKALWAYS}
  STPF_USEAPPPEEKWHENACTIVE      = $8;
  {$EXTERNALSYM STPF_USEAPPPEEKWHENACTIVE}

type
  ITaskbarList4 = interface(ITaskbarList3)
    ['{C43DC798-95D1-4BEA-9030-BB99E2983A1A}']
    function SetTabProperties(hwndTab: HWND; stpFlags: STPFLAG): HRESULT; stdcall;
  end;
  {$EXTERNALSYM ITaskbarList4}

const
  CLSID_TaskbarList: TGUID                            = '{56FDF344-FD6D-11d0-958A-006097C9A090}';
  {$EXTERNALSYM CLSID_TaskbarList}


{ TODO BCB-compatibility}

const
  DLLVER_PLATFORM_WINDOWS = $00000001;
  {$EXTERNALSYM DLLVER_PLATFORM_WINDOWS}
  DLLVER_PLATFORM_NT      = $00000002;
  {$EXTERNALSYM DLLVER_PLATFORM_NT}

type
  PDllVersionInfo = ^TDllVersionInfo;
  _DLLVERSIONINFO = packed record
    cbSize: DWORD;
    dwMajorVersion: DWORD;
    dwMinorVersion: DWORD;
    dwBuildNumber: DWORD;
    dwPlatformId: DWORD;
  end;
  {$EXTERNALSYM _DLLVERSIONINFO}
  TDllVersionInfo = _DLLVERSIONINFO;
  DLLVERSIONINFO = _DLLVERSIONINFO;
  {$EXTERNALSYM DLLVERSIONINFO}


// JwaWinError
// line 22146

const

//
// Task Scheduler errors
//
//
// MessageId: SCHED_S_TASK_READY
//
// MessageText:
//
//  The task is ready to run at its next scheduled time.
//
  SCHED_S_TASK_READY = HRESULT($00041300);
  {$EXTERNALSYM SCHED_S_TASK_READY}

//
// MessageId: SCHED_S_TASK_RUNNING
//
// MessageText:
//
//  The task is currently running.
//
  SCHED_S_TASK_RUNNING = HRESULT($00041301);
  {$EXTERNALSYM SCHED_S_TASK_RUNNING}

//
// MessageId: SCHED_S_TASK_DISABLED
//
// MessageText:
//
//  The task will not run at the scheduled times because it has been disabled.
//
  SCHED_S_TASK_DISABLED = HRESULT($00041302);
  {$EXTERNALSYM SCHED_S_TASK_DISABLED}

//
// MessageId: SCHED_S_TASK_HAS_NOT_RUN
//
// MessageText:
//
//  The task has not yet run.
//
  SCHED_S_TASK_HAS_NOT_RUN = HRESULT($00041303);
  {$EXTERNALSYM SCHED_S_TASK_HAS_NOT_RUN}

//
// MessageId: SCHED_S_TASK_NO_MORE_RUNS
//
// MessageText:
//
//  There are no more runs scheduled for this task.
//
  SCHED_S_TASK_NO_MORE_RUNS = HRESULT($00041304);
  {$EXTERNALSYM SCHED_S_TASK_NO_MORE_RUNS}

//
// MessageId: SCHED_S_TASK_NOT_SCHEDULED
//
// MessageText:
//
//  One or more of the properties that are needed to run this task on a schedule have not been set.
//
  SCHED_S_TASK_NOT_SCHEDULED = HRESULT($00041305);
  {$EXTERNALSYM SCHED_S_TASK_NOT_SCHEDULED}

//
// MessageId: SCHED_S_TASK_TERMINATED
//
// MessageText:
//
//  The last run of the task was terminated by the user.
//
  SCHED_S_TASK_TERMINATED = HRESULT($00041306);
  {$EXTERNALSYM SCHED_S_TASK_TERMINATED}

//
// MessageId: SCHED_S_TASK_NO_VALID_TRIGGERS
//
// MessageText:
//
//  Either the task has no triggers or the existing triggers are disabled or not set.
//
  SCHED_S_TASK_NO_VALID_TRIGGERS = HRESULT($00041307);
  {$EXTERNALSYM SCHED_S_TASK_NO_VALID_TRIGGERS}

//
// MessageId: SCHED_S_EVENT_TRIGGER
//
// MessageText:
//
//  Event triggers don't have set run times.
//
  SCHED_S_EVENT_TRIGGER = HRESULT($00041308);
  {$EXTERNALSYM SCHED_S_EVENT_TRIGGER}

//
// MessageId: SCHED_E_TRIGGER_NOT_FOUND
//
// MessageText:
//
//  Trigger not found.
//
  SCHED_E_TRIGGER_NOT_FOUND = HRESULT($80041309);
  {$EXTERNALSYM SCHED_E_TRIGGER_NOT_FOUND}

//
// MessageId: SCHED_E_TASK_NOT_READY
//
// MessageText:
//
//  One or more of the properties that are needed to run this task have not been set.
//
  SCHED_E_TASK_NOT_READY = HRESULT($8004130A);
  {$EXTERNALSYM SCHED_E_TASK_NOT_READY}

//
// MessageId: SCHED_E_TASK_NOT_RUNNING
//
// MessageText:
//
//  There is no running instance of the task to terminate.
//
  SCHED_E_TASK_NOT_RUNNING = HRESULT($8004130B);
  {$EXTERNALSYM SCHED_E_TASK_NOT_RUNNING}

//
// MessageId: SCHED_E_SERVICE_NOT_INSTALLED
//
// MessageText:
//
//  The Task Scheduler Service is not installed on this computer.
//
  SCHED_E_SERVICE_NOT_INSTALLED = HRESULT($8004130C);
  {$EXTERNALSYM SCHED_E_SERVICE_NOT_INSTALLED}

//
// MessageId: SCHED_E_CANNOT_OPEN_TASK
//
// MessageText:
//
//  The task object could not be opened.
//
  SCHED_E_CANNOT_OPEN_TASK = HRESULT($8004130D);
  {$EXTERNALSYM SCHED_E_CANNOT_OPEN_TASK}

//
// MessageId: SCHED_E_INVALID_TASK
//
// MessageText:
//
//  The object is either an invalid task object or is not a task object.
//
  SCHED_E_INVALID_TASK = HRESULT($8004130E);
  {$EXTERNALSYM SCHED_E_INVALID_TASK}

//
// MessageId: SCHED_E_ACCOUNT_INFORMATION_NOT_SET
//
// MessageText:
//
//  No account information could be found in the Task Scheduler security database for the task indicated.
//
  SCHED_E_ACCOUNT_INFORMATION_NOT_SET = HRESULT($8004130F);
  {$EXTERNALSYM SCHED_E_ACCOUNT_INFORMATION_NOT_SET}

//
// MessageId: SCHED_E_ACCOUNT_NAME_NOT_FOUND
//
// MessageText:
//
//  Unable to establish existence of the account specified.
//
  SCHED_E_ACCOUNT_NAME_NOT_FOUND = HRESULT($80041310);
  {$EXTERNALSYM SCHED_E_ACCOUNT_NAME_NOT_FOUND}

//
// MessageId: SCHED_E_ACCOUNT_DBASE_CORRUPT
//
// MessageText:
//
//  Corruption was detected in the Task Scheduler security database; the database has been reset.
//
  SCHED_E_ACCOUNT_DBASE_CORRUPT = HRESULT($80041311);
  {$EXTERNALSYM SCHED_E_ACCOUNT_DBASE_CORRUPT}

//
// MessageId: SCHED_E_NO_SECURITY_SERVICES
//
// MessageText:
//
//  Task Scheduler security services are available only on Windows NT.
//
  SCHED_E_NO_SECURITY_SERVICES = HRESULT($80041312);
  {$EXTERNALSYM SCHED_E_NO_SECURITY_SERVICES}

//
// MessageId: SCHED_E_UNKNOWN_OBJECT_VERSION
//
// MessageText:
//
//  The task object version is either unsupported or invalid.
//
  SCHED_E_UNKNOWN_OBJECT_VERSION = HRESULT($80041313);
  {$EXTERNALSYM SCHED_E_UNKNOWN_OBJECT_VERSION}

//
// MessageId: SCHED_E_UNSUPPORTED_ACCOUNT_OPTION
//
// MessageText:
//
//  The task has been configured with an unsupported combination of account settings and run time options.
//
  SCHED_E_UNSUPPORTED_ACCOUNT_OPTION = HRESULT($80041314);
  {$EXTERNALSYM SCHED_E_UNSUPPORTED_ACCOUNT_OPTION}

//
// MessageId: SCHED_E_SERVICE_NOT_RUNNING
//
// MessageText:
//
//  The Task Scheduler Service is not running.
//
  SCHED_E_SERVICE_NOT_RUNNING = HRESULT($80041315);
  {$EXTERNALSYM SCHED_E_SERVICE_NOT_RUNNING}


// line 151

//
// Define the various device type values.  Note that values used by Microsoft
// Corporation are in the range 0-32767, and 32768-65535 are reserved for use
// by customers.
//

type
  DEVICE_TYPE = DWORD;
  {$EXTERNALSYM DEVICE_TYPE}

const
  FILE_DEVICE_BEEP                = $00000001;
  {$EXTERNALSYM FILE_DEVICE_BEEP}
  FILE_DEVICE_CD_ROM              = $00000002;
  {$EXTERNALSYM FILE_DEVICE_CD_ROM}
  FILE_DEVICE_CD_ROM_FILE_SYSTEM  = $00000003;
  {$EXTERNALSYM FILE_DEVICE_CD_ROM_FILE_SYSTEM}
  FILE_DEVICE_CONTROLLER          = $00000004;
  {$EXTERNALSYM FILE_DEVICE_CONTROLLER}
  FILE_DEVICE_DATALINK            = $00000005;
  {$EXTERNALSYM FILE_DEVICE_DATALINK}
  FILE_DEVICE_DFS                 = $00000006;
  {$EXTERNALSYM FILE_DEVICE_DFS}
  FILE_DEVICE_DISK                = $00000007;
  {$EXTERNALSYM FILE_DEVICE_DISK}
  FILE_DEVICE_DISK_FILE_SYSTEM    = $00000008;
  {$EXTERNALSYM FILE_DEVICE_DISK_FILE_SYSTEM}
  FILE_DEVICE_FILE_SYSTEM         = $00000009;
  {$EXTERNALSYM FILE_DEVICE_FILE_SYSTEM}
  FILE_DEVICE_INPORT_PORT         = $0000000a;
  {$EXTERNALSYM FILE_DEVICE_INPORT_PORT}
  FILE_DEVICE_KEYBOARD            = $0000000b;
  {$EXTERNALSYM FILE_DEVICE_KEYBOARD}
  FILE_DEVICE_MAILSLOT            = $0000000c;
  {$EXTERNALSYM FILE_DEVICE_MAILSLOT}
  FILE_DEVICE_MIDI_IN             = $0000000d;
  {$EXTERNALSYM FILE_DEVICE_MIDI_IN}
  FILE_DEVICE_MIDI_OUT            = $0000000e;
  {$EXTERNALSYM FILE_DEVICE_MIDI_OUT}
  FILE_DEVICE_MOUSE               = $0000000f;
  {$EXTERNALSYM FILE_DEVICE_MOUSE}
  FILE_DEVICE_MULTI_UNC_PROVIDER  = $00000010;
  {$EXTERNALSYM FILE_DEVICE_MULTI_UNC_PROVIDER}
  FILE_DEVICE_NAMED_PIPE          = $00000011;
  {$EXTERNALSYM FILE_DEVICE_NAMED_PIPE}
  FILE_DEVICE_NETWORK             = $00000012;
  {$EXTERNALSYM FILE_DEVICE_NETWORK}
  FILE_DEVICE_NETWORK_BROWSER     = $00000013;
  {$EXTERNALSYM FILE_DEVICE_NETWORK_BROWSER}
  FILE_DEVICE_NETWORK_FILE_SYSTEM = $00000014;
  {$EXTERNALSYM FILE_DEVICE_NETWORK_FILE_SYSTEM}
  FILE_DEVICE_NULL                = $00000015;
  {$EXTERNALSYM FILE_DEVICE_NULL}
  FILE_DEVICE_PARALLEL_PORT       = $00000016;
  {$EXTERNALSYM FILE_DEVICE_PARALLEL_PORT}
  FILE_DEVICE_PHYSICAL_NETCARD    = $00000017;
  {$EXTERNALSYM FILE_DEVICE_PHYSICAL_NETCARD}
  FILE_DEVICE_PRINTER             = $00000018;
  {$EXTERNALSYM FILE_DEVICE_PRINTER}
  FILE_DEVICE_SCANNER             = $00000019;
  {$EXTERNALSYM FILE_DEVICE_SCANNER}
  FILE_DEVICE_SERIAL_MOUSE_PORT   = $0000001a;
  {$EXTERNALSYM FILE_DEVICE_SERIAL_MOUSE_PORT}
  FILE_DEVICE_SERIAL_PORT         = $0000001b;
  {$EXTERNALSYM FILE_DEVICE_SERIAL_PORT}
  FILE_DEVICE_SCREEN              = $0000001c;
  {$EXTERNALSYM FILE_DEVICE_SCREEN}
  FILE_DEVICE_SOUND               = $0000001d;
  {$EXTERNALSYM FILE_DEVICE_SOUND}
  FILE_DEVICE_STREAMS             = $0000001e;
  {$EXTERNALSYM FILE_DEVICE_STREAMS}
  FILE_DEVICE_TAPE                = $0000001f;
  {$EXTERNALSYM FILE_DEVICE_TAPE}
  FILE_DEVICE_TAPE_FILE_SYSTEM    = $00000020;
  {$EXTERNALSYM FILE_DEVICE_TAPE_FILE_SYSTEM}
  FILE_DEVICE_TRANSPORT           = $00000021;
  {$EXTERNALSYM FILE_DEVICE_TRANSPORT}
  FILE_DEVICE_UNKNOWN             = $00000022;
  {$EXTERNALSYM FILE_DEVICE_UNKNOWN}
  FILE_DEVICE_VIDEO               = $00000023;
  {$EXTERNALSYM FILE_DEVICE_VIDEO}
  FILE_DEVICE_VIRTUAL_DISK        = $00000024;
  {$EXTERNALSYM FILE_DEVICE_VIRTUAL_DISK}
  FILE_DEVICE_WAVE_IN             = $00000025;
  {$EXTERNALSYM FILE_DEVICE_WAVE_IN}
  FILE_DEVICE_WAVE_OUT            = $00000026;
  {$EXTERNALSYM FILE_DEVICE_WAVE_OUT}
  FILE_DEVICE_8042_PORT           = $00000027;
  {$EXTERNALSYM FILE_DEVICE_8042_PORT}
  FILE_DEVICE_NETWORK_REDIRECTOR  = $00000028;
  {$EXTERNALSYM FILE_DEVICE_NETWORK_REDIRECTOR}
  FILE_DEVICE_BATTERY             = $00000029;
  {$EXTERNALSYM FILE_DEVICE_BATTERY}
  FILE_DEVICE_BUS_EXTENDER        = $0000002a;
  {$EXTERNALSYM FILE_DEVICE_BUS_EXTENDER}
  FILE_DEVICE_MODEM               = $0000002b;
  {$EXTERNALSYM FILE_DEVICE_MODEM}
  FILE_DEVICE_VDM                 = $0000002c;
  {$EXTERNALSYM FILE_DEVICE_VDM}
  FILE_DEVICE_MASS_STORAGE        = $0000002d;
  {$EXTERNALSYM FILE_DEVICE_MASS_STORAGE}
  FILE_DEVICE_SMB                 = $0000002e;
  {$EXTERNALSYM FILE_DEVICE_SMB}
  FILE_DEVICE_KS                  = $0000002f;
  {$EXTERNALSYM FILE_DEVICE_KS}
  FILE_DEVICE_CHANGER             = $00000030;
  {$EXTERNALSYM FILE_DEVICE_CHANGER}
  FILE_DEVICE_SMARTCARD           = $00000031;
  {$EXTERNALSYM FILE_DEVICE_SMARTCARD}
  FILE_DEVICE_ACPI                = $00000032;
  {$EXTERNALSYM FILE_DEVICE_ACPI}
  FILE_DEVICE_DVD                 = $00000033;
  {$EXTERNALSYM FILE_DEVICE_DVD}
  FILE_DEVICE_FULLSCREEN_VIDEO    = $00000034;
  {$EXTERNALSYM FILE_DEVICE_FULLSCREEN_VIDEO}
  FILE_DEVICE_DFS_FILE_SYSTEM     = $00000035;
  {$EXTERNALSYM FILE_DEVICE_DFS_FILE_SYSTEM}
  FILE_DEVICE_DFS_VOLUME          = $00000036;
  {$EXTERNALSYM FILE_DEVICE_DFS_VOLUME}
  FILE_DEVICE_SERENUM             = $00000037;
  {$EXTERNALSYM FILE_DEVICE_SERENUM}
  FILE_DEVICE_TERMSRV             = $00000038;
  {$EXTERNALSYM FILE_DEVICE_TERMSRV}
  FILE_DEVICE_KSEC                = $00000039;
  {$EXTERNALSYM FILE_DEVICE_KSEC}
  FILE_DEVICE_FIPS                = $0000003A;
  {$EXTERNALSYM FILE_DEVICE_FIPS}
  FILE_DEVICE_INFINIBAND          = $0000003B;
  {$EXTERNALSYM FILE_DEVICE_INFINIBAND}

// line 297

//
// Define the method codes for how buffers are passed for I/O and FS controls
//

const
  METHOD_BUFFERED   = 0;
  {$EXTERNALSYM METHOD_BUFFERED}
  METHOD_IN_DIRECT  = 1;
  {$EXTERNALSYM METHOD_IN_DIRECT}
  METHOD_OUT_DIRECT = 2;
  {$EXTERNALSYM METHOD_OUT_DIRECT}
  METHOD_NEITHER    = 3;
  {$EXTERNALSYM METHOD_NEITHER}

//
// Define some easier to comprehend aliases:
//   METHOD_DIRECT_TO_HARDWARE (writes, aka METHOD_IN_DIRECT)
//   METHOD_DIRECT_FROM_HARDWARE (reads, aka METHOD_OUT_DIRECT)
//

  METHOD_DIRECT_TO_HARDWARE     = METHOD_IN_DIRECT;
  {$EXTERNALSYM METHOD_DIRECT_TO_HARDWARE}
  METHOD_DIRECT_FROM_HARDWARE   = METHOD_OUT_DIRECT;
  {$EXTERNALSYM METHOD_DIRECT_FROM_HARDWARE}

//
// Define the access check value for any access
//
//
// The FILE_READ_ACCESS and FILE_WRITE_ACCESS constants are also defined in
// ntioapi.h as FILE_READ_DATA and FILE_WRITE_DATA. The values for these
// constants *MUST* always be in sync.
//
//
// FILE_SPECIAL_ACCESS is checked by the NT I/O system the same as FILE_ANY_ACCESS.
// The file systems, however, may add additional access checks for I/O and FS controls
// that use this value.
//

const
  FILE_ANY_ACCESS     = 0;
  {$EXTERNALSYM FILE_ANY_ACCESS}
  FILE_SPECIAL_ACCESS = FILE_ANY_ACCESS;
  {$EXTERNALSYM FILE_SPECIAL_ACCESS}
  FILE_READ_ACCESS    = $0001;           // file & pipe
  {$EXTERNALSYM FILE_READ_ACCESS}
  FILE_WRITE_ACCESS   = $0002;           // file & pipe
  {$EXTERNALSYM FILE_WRITE_ACCESS}

// line 3425

//
// The following is a list of the native file system fsctls followed by
// additional network file system fsctls.  Some values have been
// decommissioned.
//

const

  FSCTL_REQUEST_OPLOCK_LEVEL_1 = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (0 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_REQUEST_OPLOCK_LEVEL_1}

  FSCTL_REQUEST_OPLOCK_LEVEL_2 = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (1 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_REQUEST_OPLOCK_LEVEL_2}

  FSCTL_REQUEST_BATCH_OPLOCK = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (2 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_REQUEST_BATCH_OPLOCK}

  FSCTL_OPLOCK_BREAK_ACKNOWLEDGE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (3 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_OPLOCK_BREAK_ACKNOWLEDGE}

  FSCTL_OPBATCH_ACK_CLOSE_PENDING = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (4 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_OPBATCH_ACK_CLOSE_PENDING}

  FSCTL_OPLOCK_BREAK_NOTIFY = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (5 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_OPLOCK_BREAK_NOTIFY}

  FSCTL_LOCK_VOLUME = ((FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or (6 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_LOCK_VOLUME}

  FSCTL_UNLOCK_VOLUME = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (7 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_UNLOCK_VOLUME}

  FSCTL_DISMOUNT_VOLUME = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (8 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_DISMOUNT_VOLUME}

// decommissioned fsctl value                                              9

  FSCTL_IS_VOLUME_MOUNTED = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (10 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_IS_VOLUME_MOUNTED}

  FSCTL_IS_PATHNAME_VALID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (11 shl 2) or METHOD_BUFFERED);    // PATHNAME_BUFFER,
  {$EXTERNALSYM FSCTL_IS_PATHNAME_VALID}

  FSCTL_MARK_VOLUME_DIRTY = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (12 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_MARK_VOLUME_DIRTY}

// decommissioned fsctl value                                             13

  FSCTL_QUERY_RETRIEVAL_POINTERS = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (14 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_QUERY_RETRIEVAL_POINTERS}

  FSCTL_GET_COMPRESSION = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (15 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_COMPRESSION}

  FSCTL_SET_COMPRESSION = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or ((FILE_READ_DATA or FILE_WRITE_DATA) shl 14) or
    (16 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_COMPRESSION}

// decommissioned fsctl value                                             17
// decommissioned fsctl value                                             18

  FSCTL_MARK_AS_SYSTEM_HIVE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (19 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_MARK_AS_SYSTEM_HIVE}

  FSCTL_OPLOCK_BREAK_ACK_NO_2 = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (20 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_OPLOCK_BREAK_ACK_NO_2}

  FSCTL_INVALIDATE_VOLUMES = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (21 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_INVALIDATE_VOLUMES}

  FSCTL_QUERY_FAT_BPB = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (22 shl 2) or METHOD_BUFFERED); // FSCTL_QUERY_FAT_BPB_BUFFER
  {$EXTERNALSYM FSCTL_QUERY_FAT_BPB}

  FSCTL_REQUEST_FILTER_OPLOCK = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (23 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_REQUEST_FILTER_OPLOCK}

  FSCTL_FILESYSTEM_GET_STATISTICS = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (24 shl 2) or METHOD_BUFFERED); // FILESYSTEM_STATISTICS
  {$EXTERNALSYM FSCTL_FILESYSTEM_GET_STATISTICS}

  FSCTL_GET_NTFS_VOLUME_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (25 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_NTFS_VOLUME_DATA}

  FSCTL_GET_NTFS_FILE_RECORD = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (26 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_NTFS_FILE_RECORD}

  FSCTL_GET_VOLUME_BITMAP = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (27 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_GET_VOLUME_BITMAP}

  FSCTL_GET_RETRIEVAL_POINTERS = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (28 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_GET_RETRIEVAL_POINTERS}

  FSCTL_MOVE_FILE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (29 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_MOVE_FILE}

  FSCTL_IS_VOLUME_DIRTY = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (30 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_IS_VOLUME_DIRTY}

// decomissioned fsctl value  31
(*  FSCTL_GET_HFS_INFORMATION = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (31 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_HFS_INFORMATION}
*)

  FSCTL_ALLOW_EXTENDED_DASD_IO = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (32 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_ALLOW_EXTENDED_DASD_IO}

// decommissioned fsctl value                                             33
// decommissioned fsctl value                                             34

(*
  FSCTL_READ_PROPERTY_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (33 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_READ_PROPERTY_DATA}

  FSCTL_WRITE_PROPERTY_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (34 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_WRITE_PROPERTY_DATA}
*)

  FSCTL_FIND_FILES_BY_SID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (35 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_FIND_FILES_BY_SID}

// decommissioned fsctl value                                             36
// decommissioned fsctl value                                             37

(*  FSCTL_DUMP_PROPERTY_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (37 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_DUMP_PROPERTY_DATA}
*)

  FSCTL_SET_OBJECT_ID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (38 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_OBJECT_ID}

  FSCTL_GET_OBJECT_ID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (39 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_OBJECT_ID}

  FSCTL_DELETE_OBJECT_ID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (40 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_DELETE_OBJECT_ID}

  FSCTL_SET_REPARSE_POINT = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (41 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_REPARSE_POINT}

  FSCTL_GET_REPARSE_POINT = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (42 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_GET_REPARSE_POINT}

  FSCTL_DELETE_REPARSE_POINT = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (43 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_DELETE_REPARSE_POINT}

  FSCTL_ENUM_USN_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (44 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_ENUM_USN_DATA}

  FSCTL_SECURITY_ID_CHECK = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_READ_DATA shl 14) or
    (45 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_SECURITY_ID_CHECK}

  FSCTL_READ_USN_JOURNAL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (46 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_READ_USN_JOURNAL}

  FSCTL_SET_OBJECT_ID_EXTENDED = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (47 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_OBJECT_ID_EXTENDED}

  FSCTL_CREATE_OR_GET_OBJECT_ID = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (48 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_CREATE_OR_GET_OBJECT_ID}

  FSCTL_SET_SPARSE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (49 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_SPARSE}

  FSCTL_SET_ZERO_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_WRITE_DATA shl 14) or
    (50 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SET_ZERO_DATA}

  FSCTL_QUERY_ALLOCATED_RANGES = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_READ_DATA shl 14) or
    (51 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_QUERY_ALLOCATED_RANGES}

// decommissioned fsctl value                                             52
(*
  FSCTL_ENABLE_UPGRADE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_WRITE_DATA shl 14) or
    (52 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_ENABLE_UPGRADE}
*)

  FSCTL_SET_ENCRYPTION = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (53 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_SET_ENCRYPTION}

  FSCTL_ENCRYPTION_FSCTL_IO = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (54 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_ENCRYPTION_FSCTL_IO}

  FSCTL_WRITE_RAW_ENCRYPTED = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (55 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_WRITE_RAW_ENCRYPTED}

  FSCTL_READ_RAW_ENCRYPTED = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (56 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_READ_RAW_ENCRYPTED}

  FSCTL_CREATE_USN_JOURNAL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (57 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_CREATE_USN_JOURNAL}

  FSCTL_READ_FILE_USN_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (58 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_READ_FILE_USN_DATA}

  FSCTL_WRITE_USN_CLOSE_RECORD = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (59 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_WRITE_USN_CLOSE_RECORD}

  FSCTL_EXTEND_VOLUME = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (60 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_EXTEND_VOLUME}

  FSCTL_QUERY_USN_JOURNAL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (61 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_QUERY_USN_JOURNAL}

  FSCTL_DELETE_USN_JOURNAL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (62 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_DELETE_USN_JOURNAL}

  FSCTL_MARK_HANDLE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (63 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_MARK_HANDLE}

  FSCTL_SIS_COPYFILE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (64 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SIS_COPYFILE}

  FSCTL_SIS_LINK_FILES = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or ((FILE_READ_DATA or FILE_WRITE_DATA) shl 14) or
    (65 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_SIS_LINK_FILES}

  FSCTL_HSM_MSG = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or ((FILE_READ_DATA or FILE_WRITE_DATA) shl 14) or
    (66 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_HSM_MSG}

// decommissioned fsctl value                                             67
(*
  FSCTL_NSS_CONTROL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_WRITE_DATA shl 14) or
    (67 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_NSS_CONTROL}
*)

  FSCTL_HSM_DATA = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or ((FILE_READ_DATA or FILE_WRITE_DATA) shl 14) or
    (68 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_HSM_DATA}

  FSCTL_RECALL_FILE = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_ANY_ACCESS shl 14) or
    (69 shl 2) or METHOD_NEITHER);
  {$EXTERNALSYM FSCTL_RECALL_FILE}

// decommissioned fsctl value                                             70
(*
  FSCTL_NSS_RCONTROL = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_READ_DATA shl 14) or
    (70 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_NSS_RCONTROL}
*)

  FSCTL_READ_FROM_PLEX = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_READ_DATA shl 14) or
    (71 shl 2) or METHOD_OUT_DIRECT);
  {$EXTERNALSYM FSCTL_READ_FROM_PLEX}

  FSCTL_FILE_PREFETCH = (
    (FILE_DEVICE_FILE_SYSTEM shl 16) or (FILE_SPECIAL_ACCESS shl 14) or
    (72 shl 2) or METHOD_BUFFERED);
  {$EXTERNALSYM FSCTL_FILE_PREFETCH}

// line 4553

//
// Structure for FSCTL_SET_ZERO_DATA
//

type

  PFILE_ZERO_DATA_INFORMATION = ^FILE_ZERO_DATA_INFORMATION;
  {$EXTERNALSYM PFILE_ZERO_DATA_INFORMATION}
  _FILE_ZERO_DATA_INFORMATION = record
    FileOffset: LARGE_INTEGER;
    BeyondFinalZero: LARGE_INTEGER;
  end;
  {$EXTERNALSYM _FILE_ZERO_DATA_INFORMATION}
  FILE_ZERO_DATA_INFORMATION = _FILE_ZERO_DATA_INFORMATION;
  {$EXTERNALSYM FILE_ZERO_DATA_INFORMATION}
  TFileZeroDataInformation = FILE_ZERO_DATA_INFORMATION;
  PFileZeroDataInformation = PFILE_ZERO_DATA_INFORMATION;

//
// Structure for FSCTL_QUERY_ALLOCATED_RANGES
//

//
// Querying the allocated ranges requires an output buffer to store the
// allocated ranges and an input buffer to specify the range to query.
// The input buffer contains a single entry, the output buffer is an
// array of the following structure.
//

  PFILE_ALLOCATED_RANGE_BUFFER = ^FILE_ALLOCATED_RANGE_BUFFER;
  {$EXTERNALSYM PFILE_ALLOCATED_RANGE_BUFFER}
  _FILE_ALLOCATED_RANGE_BUFFER = record
    FileOffset: LARGE_INTEGER;
    Length: LARGE_INTEGER;
  end;
  {$EXTERNALSYM _FILE_ALLOCATED_RANGE_BUFFER}
  FILE_ALLOCATED_RANGE_BUFFER = _FILE_ALLOCATED_RANGE_BUFFER;
  {$EXTERNALSYM FILE_ALLOCATED_RANGE_BUFFER}
  TFileAllocatedRangeBuffer = FILE_ALLOCATED_RANGE_BUFFER;
  PFileAllocatedRangeBuffer = PFILE_ALLOCATED_RANGE_BUFFER;


// line 340

//
//  Code Page Default Values.
//

const
  CP_ACP        = 0; // default to ANSI code page
  {$EXTERNALSYM CP_ACP}
  CP_OEMCP      = 1; // default to OEM  code page
  {$EXTERNALSYM CP_OEMCP}
  CP_MACCP      = 2; // default to MAC  code page
  {$EXTERNALSYM CP_MACCP}
  CP_THREAD_ACP = 3; // current thread's ANSI code page
  {$EXTERNALSYM CP_THREAD_ACP}
  CP_SYMBOL     = 42; // SYMBOL translations
  {$EXTERNALSYM CP_SYMBOL}

  CP_UTF7 = 65000; // UTF-7 translation
  {$EXTERNALSYM CP_UTF7}
  CP_UTF8 = 65001; // UTF-8 translation
  {$EXTERNALSYM CP_UTF8}

// line 597

const

//
//  The following LCTypes may be used in combination with any other LCTypes.
//
//    LOCALE_NOUSEROVERRIDE is also used in GetTimeFormat and
//    GetDateFormat.
//
//    LOCALE_USE_CP_ACP is used in many of the A (Ansi) apis that need
//    to do string translation.
//
//    LOCALE_RETURN_NUMBER will return the result from GetLocaleInfo as a
//    number instead of a string.  This flag is only valid for the LCTypes
//    beginning with LOCALE_I.
//

  LOCALE_NOUSEROVERRIDE = DWORD($80000000); // do not use user overrides
  {$EXTERNALSYM LOCALE_NOUSEROVERRIDE}
  LOCALE_USE_CP_ACP     = $40000000; // use the system ACP
  {$EXTERNALSYM LOCALE_USE_CP_ACP}

  LOCALE_RETURN_NUMBER = $20000000; // return number instead of string
  {$EXTERNALSYM LOCALE_RETURN_NUMBER}

// line 841

const
  LOCALE_IDEFAULTEBCDICCODEPAGE = $00001012; // default ebcdic code page
  {$EXTERNALSYM LOCALE_IDEFAULTEBCDICCODEPAGE}
  LOCALE_IPAPERSIZE             = $0000100A; // 1 = letter, 5 = legal, 8 = a3, 9 = a4
  {$EXTERNALSYM LOCALE_IPAPERSIZE}
  LOCALE_SENGCURRNAME           = $00001007; // english name of currency
  {$EXTERNALSYM LOCALE_SENGCURRNAME}
  LOCALE_SNATIVECURRNAME        = $00001008; // native name of currency
  {$EXTERNALSYM LOCALE_SNATIVECURRNAME}
  LOCALE_SYEARMONTH             = $00001006; // year month format string
  {$EXTERNALSYM LOCALE_SYEARMONTH}
  LOCALE_SSORTNAME              = $00001013; // sort name
  {$EXTERNALSYM LOCALE_SSORTNAME}
  LOCALE_IDIGITSUBSTITUTION     = $00001014; // 0 = context, 1 = none, 2 = national
  {$EXTERNALSYM LOCALE_IDIGITSUBSTITUTION}

// line 880

  DATE_YEARMONTH  = $00000008; // use year month picture
  {$EXTERNALSYM DATE_YEARMONTH}
  DATE_LTRREADING = $00000010; // add marks for left to right reading order layout
  {$EXTERNALSYM DATE_LTRREADING}
  DATE_RTLREADING = $00000020; // add marks for right to left reading order layout
  {$EXTERNALSYM DATE_RTLREADING}

//
//  Calendar Types.
//
//  These types are used for the EnumCalendarInfo and GetCalendarInfo
//  NLS API routines.
//  Some of these types are also used for the SetCalendarInfo NLS API
//  routine.
//

//
//  The following CalTypes may be used in combination with any other CalTypes.
//
//    CAL_NOUSEROVERRIDE
//
//    CAL_USE_CP_ACP is used in the A (Ansi) apis that need to do string
//    translation.
//
//    CAL_RETURN_NUMBER will return the result from GetCalendarInfo as a
//    number instead of a string.  This flag is only valid for the CalTypes
//    beginning with CAL_I.
//

  CAL_NOUSEROVERRIDE = LOCALE_NOUSEROVERRIDE; // do not use user overrides
  {$EXTERNALSYM CAL_NOUSEROVERRIDE}
  CAL_USE_CP_ACP     = LOCALE_USE_CP_ACP; // use the system ACP
  {$EXTERNALSYM CAL_USE_CP_ACP}
  CAL_RETURN_NUMBER  = LOCALE_RETURN_NUMBER; // return number instead of string
  {$EXTERNALSYM CAL_RETURN_NUMBER}

// line 1014

  CAL_SYEARMONTH       = $0000002f; // year month format string
  {$EXTERNALSYM CAL_SYEARMONTH}
  CAL_ITWODIGITYEARMAX = $00000030; // two digit year max
  {$EXTERNALSYM CAL_ITWODIGITYEARMAX}

// line 1424

type
  CALINFO_ENUMPROCEXW = function (lpCalendarInfoString: LPWSTR; Calendar: CALID): BOOL; stdcall;
  {$EXTERNALSYM CALINFO_ENUMPROCEXW}
  TCalInfoEnumProcExW = CALINFO_ENUMPROCEXW;


type
  MAKEINTRESOURCEA = LPSTR;
  {$EXTERNALSYM MAKEINTRESOURCEA}
  MAKEINTRESOURCEW = LPWSTR;
  {$EXTERNALSYM MAKEINTRESOURCEW}
  MAKEINTRESOURCE = MAKEINTRESOURCEW;
  {$EXTERNALSYM MAKEINTRESOURCE}

//
// Predefined Resource Types
//

const
  RT_CURSOR       = MAKEINTRESOURCE(1);
  {$EXTERNALSYM RT_CURSOR}
  RT_BITMAP       = MAKEINTRESOURCE(2);
  {$EXTERNALSYM RT_BITMAP}
  RT_ICON         = MAKEINTRESOURCE(3);
  {$EXTERNALSYM RT_ICON}
  RT_MENU         = MAKEINTRESOURCE(4);
  {$EXTERNALSYM RT_MENU}
  RT_DIALOG       = MAKEINTRESOURCE(5);
  {$EXTERNALSYM RT_DIALOG}
  RT_STRING       = MAKEINTRESOURCE(6);
  {$EXTERNALSYM RT_STRING}
  RT_FONTDIR      = MAKEINTRESOURCE(7);
  {$EXTERNALSYM RT_FONTDIR}
  RT_FONT         = MAKEINTRESOURCE(8);
  {$EXTERNALSYM RT_FONT}
  RT_ACCELERATOR  = MAKEINTRESOURCE(9);
  {$EXTERNALSYM RT_ACCELERATOR}
  RT_RCDATA       = MAKEINTRESOURCE(10);
  {$EXTERNALSYM RT_RCDATA}
  RT_MESSAGETABLE = MAKEINTRESOURCE(11);
  {$EXTERNALSYM RT_MESSAGETABLE}

  DIFFERENCE = 11;
  {$EXTERNALSYM DIFFERENCE}

  RT_GROUP_CURSOR = MAKEINTRESOURCE(ULONG_PTR(RT_CURSOR) + DIFFERENCE);
  {$EXTERNALSYM RT_GROUP_CURSOR}
  RT_GROUP_ICON = MAKEINTRESOURCE(ULONG_PTR(RT_ICON) + DIFFERENCE);
  {$EXTERNALSYM RT_GROUP_ICON}
  RT_VERSION    = MAKEINTRESOURCE(16);
  {$EXTERNALSYM RT_VERSION}
  RT_DLGINCLUDE = MAKEINTRESOURCE(17);
  {$EXTERNALSYM RT_DLGINCLUDE}
  RT_PLUGPLAY   = MAKEINTRESOURCE(19);
  {$EXTERNALSYM RT_PLUGPLAY}
  RT_VXD        = MAKEINTRESOURCE(20);
  {$EXTERNALSYM RT_VXD}
  RT_ANICURSOR  = MAKEINTRESOURCE(21);
  {$EXTERNALSYM RT_ANICURSOR}
  RT_ANIICON    = MAKEINTRESOURCE(22);
  {$EXTERNALSYM RT_ANIICON}
  RT_HTML       = MAKEINTRESOURCE(23);
  {$EXTERNALSYM RT_HTML}
  RT_MANIFEST   = MAKEINTRESOURCE(24);
  CREATEPROCESS_MANIFEST_RESOURCE_ID = MAKEINTRESOURCE(1);
  {$EXTERNALSYM CREATEPROCESS_MANIFEST_RESOURCE_ID}
  ISOLATIONAWARE_MANIFEST_RESOURCE_ID = MAKEINTRESOURCE(2);
  {$EXTERNALSYM ISOLATIONAWARE_MANIFEST_RESOURCE_ID}
  ISOLATIONAWARE_NOSTATICIMPORT_MANIFEST_RESOURCE_ID = MAKEINTRESOURCE(3);
  {$EXTERNALSYM ISOLATIONAWARE_NOSTATICIMPORT_MANIFEST_RESOURCE_ID}
  MINIMUM_RESERVED_MANIFEST_RESOURCE_ID = MAKEINTRESOURCE(1{inclusive});
  {$EXTERNALSYM MINIMUM_RESERVED_MANIFEST_RESOURCE_ID}
  MAXIMUM_RESERVED_MANIFEST_RESOURCE_ID = MAKEINTRESOURCE(16{inclusive});
  {$EXTERNALSYM MAXIMUM_RESERVED_MANIFEST_RESOURCE_ID}

// line 1451

  KLF_SETFORPROCESS = $00000100;
  {$EXTERNALSYM KLF_SETFORPROCESS}
  KLF_SHIFTLOCK     = $00010000;
  {$EXTERNALSYM KLF_SHIFTLOCK}
  KLF_RESET         = $40000000;
  {$EXTERNALSYM KLF_RESET}

// 64 compatible version of GetWindowLong and SetWindowLong

const
  GWLP_WNDPROC    = -4;
  {$EXTERNALSYM GWLP_WNDPROC}
  GWLP_HINSTANCE  = -6;
  {$EXTERNALSYM GWLP_HINSTANCE}
  GWLP_HWNDPARENT = -8;
  {$EXTERNALSYM GWLP_HWNDPARENT}
  GWLP_USERDATA   = -21;
  {$EXTERNALSYM GWLP_USERDATA}
  GWLP_ID         = -12;
  {$EXTERNALSYM GWLP_ID}


type
  // Microsoft version (64 bit SDK)
  {$EXTERNALSYM RVA}
  RVA = DWORD;

  // 64-bit PE
  {$EXTERNALSYM ImgDelayDescrV2}
  ImgDelayDescrV2 = packed record
    grAttrs: DWORD;      // attributes
    rvaDLLName: RVA;     // RVA to dll name
    rvaHmod: RVA;        // RVA of module handle
    rvaIAT: RVA;         // RVA of the IAT
    rvaINT: RVA;         // RVA of the INT
    rvaBoundIAT: RVA;    // RVA of the optional bound IAT
    rvaUnloadIAT: RVA;   // RVA of optional copy of original IAT
    dwTimeStamp: DWORD;  // 0 if not bound,
                         // O.W. date/time stamp of DLL bound to (Old BIND)
  end;
  {$EXTERNALSYM TImgDelayDescrV2}
  TImgDelayDescrV2 = ImgDelayDescrV2;
  {$EXTERNALSYM PImgDelayDescrV2}
  PImgDelayDescrV2 = ^ImgDelayDescrV2;

  {$EXTERNALSYM PHMODULE}
  PHMODULE = ^HMODULE;

  // 32-bit PE
  {$EXTERNALSYM ImgDelayDescrV1}
  ImgDelayDescrV1 = packed record
    grAttrs: DWORD;                // attributes
    szName: LPCSTR;                // pointer to dll name
    phmod: PHMODULE;               // address of module handle
    pIAT: PImageThunkData32;       // address of the IAT
    pINT: PImageThunkData32;       // address of the INT
    pBoundIAT: PImageThunkData32;  // address of the optional bound IAT
    pUnloadIAT: PImageThunkData32; // address of optional copy of original IAT
    dwTimeStamp: DWORD;            // 0 if not bound,
                                   // O.W. date/time stamp of DLL bound to (Old BIND)
  end;
  {$EXTERNALSYM TImgDelayDescrV1}
  TImgDelayDescrV1 = ImgDelayDescrV1;
  {$EXTERNALSYM PImgDelayDescrV1}
  PImgDelayDescrV1 = ^ImgDelayDescrV1;

  //{$EXTERNALSYM PImgDelayDescr}
  //PImgDelayDescr = ImgDelayDescr;
  //TImgDelayDescr = ImgDelayDescr;

// msidefs.h line 349

// PIDs given specific meanings for Installer

const
  PID_MSIVERSION  = $0000000E; // integer, Installer version number (major*100+minor)
  {$EXTERNALSYM PID_MSIVERSION}
  PID_MSISOURCE   = $0000000F; // integer, type of file image, short/long, media/tree
  {$EXTERNALSYM PID_MSISOURCE}
  PID_MSIRESTRICT = $00000010; // integer, transform restrictions
  {$EXTERNALSYM PID_MSIRESTRICT}


// shlguid.h line 404

const
  FMTID_ShellDetails: TGUID = '{28636aa6-953d-11d2-b5d6-00c04fd918d0}';
  {$EXTERNALSYM FMTID_ShellDetails}

  PID_FINDDATA        = 0;
  {$EXTERNALSYM PID_FINDDATA}
  PID_NETRESOURCE     = 1;
  {$EXTERNALSYM PID_NETRESOURCE}
  PID_DESCRIPTIONID   = 2;
  {$EXTERNALSYM PID_DESCRIPTIONID}
  PID_WHICHFOLDER     = 3;
  {$EXTERNALSYM PID_WHICHFOLDER}
  PID_NETWORKLOCATION = 4;
  {$EXTERNALSYM PID_NETWORKLOCATION}
  PID_COMPUTERNAME    = 5;
  {$EXTERNALSYM PID_COMPUTERNAME}

// PSGUID_STORAGE comes from ntquery.h
const
  FMTID_Storage: TGUID = '{b725f130-47ef-101a-a5f1-02608c9eebac}';
  {$EXTERNALSYM FMTID_Storage}

// Image properties
const
  FMTID_ImageProperties: TGUID = '{14b81da1-0135-4d31-96d9-6cbfc9671a99}';
  {$EXTERNALSYM FMTID_ImageProperties}

// The GUIDs used to identify shell item attributes (columns). See IShellFolder2::GetDetailsEx implementations...

const
  FMTID_Displaced: TGUID = '{9B174B33-40FF-11d2-A27E-00C04FC30871}';
  {$EXTERNALSYM FMTID_Displaced}
  PID_DISPLACED_FROM = 2;
  {$EXTERNALSYM PID_DISPLACED_FROM}
  PID_DISPLACED_DATE = 3;
  {$EXTERNALSYM PID_DISPLACED_DATE}

const
  FMTID_Briefcase: TGUID = '{328D8B21-7729-4bfc-954C-902B329D56B0}';
  {$EXTERNALSYM FMTID_Briefcase}
  PID_SYNC_COPY_IN = 2;
  {$EXTERNALSYM PID_SYNC_COPY_IN}

const
  FMTID_Misc: TGUID = '{9B174B34-40FF-11d2-A27E-00C04FC30871}';
  {$EXTERNALSYM FMTID_Misc}
  PID_MISC_STATUS      = 2;
  {$EXTERNALSYM PID_MISC_STATUS}
  PID_MISC_ACCESSCOUNT = 3;
  {$EXTERNALSYM PID_MISC_ACCESSCOUNT}
  PID_MISC_OWNER       = 4;
  {$EXTERNALSYM PID_MISC_OWNER}
  PID_HTMLINFOTIPFILE  = 5;
  {$EXTERNALSYM PID_HTMLINFOTIPFILE}
  PID_MISC_PICS        = 6;
  {$EXTERNALSYM PID_MISC_PICS}

const
  FMTID_WebView: TGUID = '{F2275480-F782-4291-BD94-F13693513AEC}';
  {$EXTERNALSYM FMTID_WebView}
  PID_DISPLAY_PROPERTIES = 0;
  {$EXTERNALSYM PID_DISPLAY_PROPERTIES}
  PID_INTROTEXT          = 1;
  {$EXTERNALSYM PID_INTROTEXT}

const
  FMTID_MUSIC: TGUID = '{56A3372E-CE9C-11d2-9F0E-006097C686F6}';
  {$EXTERNALSYM FMTID_MUSIC}
  PIDSI_ARTIST    = 2;
  {$EXTERNALSYM PIDSI_ARTIST}
  PIDSI_SONGTITLE = 3;
  {$EXTERNALSYM PIDSI_SONGTITLE}
  PIDSI_ALBUM     = 4;
  {$EXTERNALSYM PIDSI_ALBUM}
  PIDSI_YEAR      = 5;
  {$EXTERNALSYM PIDSI_YEAR}
  PIDSI_COMMENT   = 6;
  {$EXTERNALSYM PIDSI_COMMENT}
  PIDSI_TRACK     = 7;
  {$EXTERNALSYM PIDSI_TRACK}
  PIDSI_GENRE     = 11;
  {$EXTERNALSYM PIDSI_GENRE}
  PIDSI_LYRICS    = 12;
  {$EXTERNALSYM PIDSI_LYRICS}

const
  FMTID_DRM: TGUID = '{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}';
  {$EXTERNALSYM FMTID_DRM}
  PIDDRSI_PROTECTED   = 2;
  {$EXTERNALSYM PIDDRSI_PROTECTED}
  PIDDRSI_DESCRIPTION = 3;
  {$EXTERNALSYM PIDDRSI_DESCRIPTION}
  PIDDRSI_PLAYCOUNT   = 4;
  {$EXTERNALSYM PIDDRSI_PLAYCOUNT}
  PIDDRSI_PLAYSTARTS  = 5;
  {$EXTERNALSYM PIDDRSI_PLAYSTARTS}
  PIDDRSI_PLAYEXPIRES = 6;
  {$EXTERNALSYM PIDDRSI_PLAYEXPIRES}

//  FMTID_VideoSummaryInformation property identifiers
const
  FMTID_Video: TGUID = '{64440491-4c8b-11d1-8b70-080036b11a03}';
  {$EXTERNALSYM FMTID_Video}
  PIDVSI_STREAM_NAME   = $00000002; // "StreamName", VT_LPWSTR
  {$EXTERNALSYM PIDVSI_STREAM_NAME}
  PIDVSI_FRAME_WIDTH   = $00000003; // "FrameWidth", VT_UI4
  {$EXTERNALSYM PIDVSI_FRAME_WIDTH}
  PIDVSI_FRAME_HEIGHT  = $00000004; // "FrameHeight", VT_UI4
  {$EXTERNALSYM PIDVSI_FRAME_HEIGHT}
  PIDVSI_TIMELENGTH    = $00000007; // "TimeLength", VT_UI4, milliseconds
  {$EXTERNALSYM PIDVSI_TIMELENGTH}
  PIDVSI_FRAME_COUNT   = $00000005; // "FrameCount". VT_UI4
  {$EXTERNALSYM PIDVSI_FRAME_COUNT}
  PIDVSI_FRAME_RATE    = $00000006; // "FrameRate", VT_UI4, frames/millisecond
  {$EXTERNALSYM PIDVSI_FRAME_RATE}
  PIDVSI_DATA_RATE     = $00000008; // "DataRate", VT_UI4, bytes/second
  {$EXTERNALSYM PIDVSI_DATA_RATE}
  PIDVSI_SAMPLE_SIZE   = $00000009; // "SampleSize", VT_UI4
  {$EXTERNALSYM PIDVSI_SAMPLE_SIZE}
  PIDVSI_COMPRESSION   = $0000000A; // "Compression", VT_LPWSTR
  {$EXTERNALSYM PIDVSI_COMPRESSION}
  PIDVSI_STREAM_NUMBER = $0000000B; // "StreamNumber", VT_UI2
  {$EXTERNALSYM PIDVSI_STREAM_NUMBER}

//  FMTID_AudioSummaryInformation property identifiers
const
  FMTID_Audio: TGUID = '{64440490-4c8b-11d1-8b70-080036b11a03}';
  {$EXTERNALSYM FMTID_Audio}
  PIDASI_FORMAT        = $00000002; // VT_BSTR
  {$EXTERNALSYM PIDASI_FORMAT}
  PIDASI_TIMELENGTH    = $00000003; // VT_UI4, milliseconds
  {$EXTERNALSYM PIDASI_TIMELENGTH}
  PIDASI_AVG_DATA_RATE = $00000004; // VT_UI4,  Hz
  {$EXTERNALSYM PIDASI_AVG_DATA_RATE}
  PIDASI_SAMPLE_RATE   = $00000005; // VT_UI4,  bits
  {$EXTERNALSYM PIDASI_SAMPLE_RATE}
  PIDASI_SAMPLE_SIZE   = $00000006; // VT_UI4,  bits
  {$EXTERNALSYM PIDASI_SAMPLE_SIZE}
  PIDASI_CHANNEL_COUNT = $00000007; // VT_UI4
  {$EXTERNALSYM PIDASI_CHANNEL_COUNT}
  PIDASI_STREAM_NUMBER = $00000008; // VT_UI2
  {$EXTERNALSYM PIDASI_STREAM_NUMBER}
  PIDASI_STREAM_NAME   = $00000009; // VT_LPWSTR
  {$EXTERNALSYM PIDASI_STREAM_NAME}
  PIDASI_COMPRESSION   = $0000000A; // VT_LPWSTR
  {$EXTERNALSYM PIDASI_COMPRESSION}

const
  FMTID_ControlPanel: TGUID = '{305CA226-D286-468e-B848-2B2E8E697B74}';
  {$EXTERNALSYM FMTID_ControlPanel}
  PID_CONTROLPANEL_CATEGORY = 2;
  {$EXTERNALSYM PID_CONTROLPANEL_CATEGORY}

const
  FMTID_Volume: TGUID = '{9B174B35-40FF-11d2-A27E-00C04FC30871}';
  {$EXTERNALSYM FMTID_Volume}
  PID_VOLUME_FREE       = 2;
  {$EXTERNALSYM PID_VOLUME_FREE}
  PID_VOLUME_CAPACITY   = 3;
  {$EXTERNALSYM PID_VOLUME_CAPACITY}
  PID_VOLUME_FILESYSTEM = 4;
  {$EXTERNALSYM PID_VOLUME_FILESYSTEM}

const
  FMTID_Share: TGUID = '{D8C3986F-813B-449c-845D-87B95D674ADE}';
  {$EXTERNALSYM FMTID_Share}
  PID_SHARE_CSC_STATUS = 2;
  {$EXTERNALSYM PID_SHARE_CSC_STATUS}

const
  FMTID_Link: TGUID = '{B9B4B3FC-2B51-4a42-B5D8-324146AFCF25}';
  {$EXTERNALSYM FMTID_Link}
  PID_LINK_TARGET = 2;
  {$EXTERNALSYM PID_LINK_TARGET}

const
  FMTID_Query: TGUID = '{49691c90-7e17-101a-a91c-08002b2ecda9}';
  {$EXTERNALSYM FMTID_Query}
  PID_QUERY_RANK = 2;
  {$EXTERNALSYM PID_QUERY_RANK}

const
  FMTID_SummaryInformation: TGUID = '{f29f85e0-4ff9-1068-ab91-08002b27b3d9}';
  {$EXTERNALSYM FMTID_SummaryInformation}
  FMTID_DocumentSummaryInformation: TGUID = '{d5cdd502-2e9c-101b-9397-08002b2cf9ae}';
  {$EXTERNALSYM FMTID_DocumentSummaryInformation}
  FMTID_MediaFileSummaryInformation: TGUID = '{64440492-4c8b-11d1-8b70-080036b11a03}';
  {$EXTERNALSYM FMTID_MediaFileSummaryInformation}
  FMTID_ImageSummaryInformation: TGUID = '{6444048f-4c8b-11d1-8b70-080036b11a03}';
  {$EXTERNALSYM FMTID_ImageSummaryInformation}

// imgguids.h line 75

// Property sets
const
  FMTID_ImageInformation: TGUID = '{e5836cbe-5eef-4f1d-acde-ae4c43b608ce}';
  {$EXTERNALSYM FMTID_ImageInformation}
  FMTID_JpegAppHeaders: TGUID = '{1c4afdcd-6177-43cf-abc7-5f51af39ee85}';
  {$EXTERNALSYM FMTID_JpegAppHeaders}



// objbase.h line 390
const
  STGFMT_STORAGE  = 0;
  {$EXTERNALSYM STGFMT_STORAGE}
  STGFMT_NATIVE   = 1;
  {$EXTERNALSYM STGFMT_NATIVE}
  STGFMT_FILE     = 3;
  {$EXTERNALSYM STGFMT_FILE}
  STGFMT_ANY      = 4;
  {$EXTERNALSYM STGFMT_ANY}
  STGFMT_DOCFILE  = 5;
  {$EXTERNALSYM STGFMT_DOCFILE}
// This is a legacy define to allow old component to builds
  STGFMT_DOCUMENT = 0;
  {$EXTERNALSYM STGFMT_DOCUMENT}

// objbase.h line 913

type
  tagSTGOPTIONS = record
    usVersion: Word;             // Versions 1 and 2 supported
    reserved: Word;              // must be 0 for padding
    ulSectorSize: Cardinal;      // docfile header sector size (512)
    pwcsTemplateFile: PWideChar; // version 2 or above
  end;
  {$EXTERNALSYM tagSTGOPTIONS}
  STGOPTIONS = tagSTGOPTIONS;
  {$EXTERNALSYM STGOPTIONS}
  PSTGOPTIONS = ^STGOPTIONS;
  {$EXTERNALSYM PSTGOPTIONS}


// propidl.h line 386

// Reserved global Property IDs
const
  PID_DICTIONARY         = $00000000; // integer count + array of entries
  {$EXTERNALSYM PID_DICTIONARY}
  PID_CODEPAGE           = $00000001; // short integer
  {$EXTERNALSYM PID_CODEPAGE}
  PID_FIRST_USABLE       = $00000002;
  {$EXTERNALSYM PID_FIRST_USABLE}
  PID_FIRST_NAME_DEFAULT = $00000FFF;
  {$EXTERNALSYM PID_FIRST_NAME_DEFAULT}
  PID_LOCALE             = $80000000;
  {$EXTERNALSYM PID_LOCALE}
  PID_MODIFY_TIME        = $80000001;
  {$EXTERNALSYM PID_MODIFY_TIME}
  PID_SECURITY           = $80000002;
  {$EXTERNALSYM PID_SECURITY}
  PID_BEHAVIOR           = $80000003;
  {$EXTERNALSYM PID_BEHAVIOR}
  PID_ILLEGAL            = $FFFFFFFF;
  {$EXTERNALSYM PID_ILLEGAL}

// Range which is read-only to downlevel implementations

const
  PID_MIN_READONLY = $80000000;
  {$EXTERNALSYM PID_MIN_READONLY}
  PID_MAX_READONLY = $BFFFFFFF;
  {$EXTERNALSYM PID_MAX_READONLY}

// Property IDs for the DiscardableInformation Property Set

const
  PIDDI_THUMBNAIL = $00000002; // VT_BLOB
  {$EXTERNALSYM PIDDI_THUMBNAIL}

// Property IDs for the SummaryInformation Property Set

const
  PIDSI_TITLE        = $00000002; // VT_LPSTR
  {$EXTERNALSYM PIDSI_TITLE}
  PIDSI_SUBJECT      = $00000003; // VT_LPSTR
  {$EXTERNALSYM PIDSI_SUBJECT}
  PIDSI_AUTHOR       = $00000004; // VT_LPSTR
  {$EXTERNALSYM PIDSI_AUTHOR}
  PIDSI_KEYWORDS     = $00000005; // VT_LPSTR
  {$EXTERNALSYM PIDSI_KEYWORDS}
  PIDSI_COMMENTS     = $00000006; // VT_LPSTR
  {$EXTERNALSYM PIDSI_COMMENTS}
  PIDSI_TEMPLATE     = $00000007; // VT_LPSTR
  {$EXTERNALSYM PIDSI_TEMPLATE}
  PIDSI_LASTAUTHOR   = $00000008; // VT_LPSTR
  {$EXTERNALSYM PIDSI_LASTAUTHOR}
  PIDSI_REVNUMBER    = $00000009; // VT_LPSTR
  {$EXTERNALSYM PIDSI_REVNUMBER}
  PIDSI_EDITTIME     = $0000000A; // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_EDITTIME}
  PIDSI_LASTPRINTED  = $0000000B; // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_LASTPRINTED}
  PIDSI_CREATE_DTM   = $0000000C; // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_CREATE_DTM}
  PIDSI_LASTSAVE_DTM = $0000000D; // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDSI_LASTSAVE_DTM}
  PIDSI_PAGECOUNT    = $0000000E; // VT_I4
  {$EXTERNALSYM PIDSI_PAGECOUNT}
  PIDSI_WORDCOUNT    = $0000000F; // VT_I4
  {$EXTERNALSYM PIDSI_WORDCOUNT}
  PIDSI_CHARCOUNT    = $00000010; // VT_I4
  {$EXTERNALSYM PIDSI_CHARCOUNT}
  PIDSI_THUMBNAIL    = $00000011; // VT_CF
  {$EXTERNALSYM PIDSI_THUMBNAIL}
  PIDSI_APPNAME      = $00000012; // VT_LPSTR
  {$EXTERNALSYM PIDSI_APPNAME}
  PIDSI_DOC_SECURITY = $00000013; // VT_I4
  {$EXTERNALSYM PIDSI_DOC_SECURITY}

// Property IDs for the DocSummaryInformation Property Set

const
  PIDDSI_CATEGORY    = $00000002; // VT_LPSTR
  {$EXTERNALSYM PIDDSI_CATEGORY}
  PIDDSI_PRESFORMAT  = $00000003; // VT_LPSTR
  {$EXTERNALSYM PIDDSI_PRESFORMAT}
  PIDDSI_BYTECOUNT   = $00000004; // VT_I4
  {$EXTERNALSYM PIDDSI_BYTECOUNT}
  PIDDSI_LINECOUNT   = $00000005; // VT_I4
  {$EXTERNALSYM PIDDSI_LINECOUNT}
  PIDDSI_PARCOUNT    = $00000006; // VT_I4
  {$EXTERNALSYM PIDDSI_PARCOUNT}
  PIDDSI_SLIDECOUNT  = $00000007; // VT_I4
  {$EXTERNALSYM PIDDSI_SLIDECOUNT}
  PIDDSI_NOTECOUNT   = $00000008; // VT_I4
  {$EXTERNALSYM PIDDSI_NOTECOUNT}
  PIDDSI_HIDDENCOUNT = $00000009; // VT_I4
  {$EXTERNALSYM PIDDSI_HIDDENCOUNT}
  PIDDSI_MMCLIPCOUNT = $0000000A; // VT_I4
  {$EXTERNALSYM PIDDSI_MMCLIPCOUNT}
  PIDDSI_SCALE       = $0000000B; // VT_BOOL
  {$EXTERNALSYM PIDDSI_SCALE}
  PIDDSI_HEADINGPAIR = $0000000C; // VT_VARIANT | VT_VECTOR
  {$EXTERNALSYM PIDDSI_HEADINGPAIR}
  PIDDSI_DOCPARTS    = $0000000D; // VT_LPSTR | VT_VECTOR
  {$EXTERNALSYM PIDDSI_DOCPARTS}
  PIDDSI_MANAGER     = $0000000E; // VT_LPSTR
  {$EXTERNALSYM PIDDSI_MANAGER}
  PIDDSI_COMPANY     = $0000000F; // VT_LPSTR
  {$EXTERNALSYM PIDDSI_COMPANY}
  PIDDSI_LINKSDIRTY  = $00000010; // VT_BOOL
  {$EXTERNALSYM PIDDSI_LINKSDIRTY}

//  FMTID_MediaFileSummaryInfo - Property IDs

const
  PIDMSI_EDITOR      = $00000002; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_EDITOR}
  PIDMSI_SUPPLIER    = $00000003; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_SUPPLIER}
  PIDMSI_SOURCE      = $00000004; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_SOURCE}
  PIDMSI_SEQUENCE_NO = $00000005; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_SEQUENCE_NO}
  PIDMSI_PROJECT     = $00000006; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_PROJECT}
  PIDMSI_STATUS      = $00000007; // VT_UI4
  {$EXTERNALSYM PIDMSI_STATUS}
  PIDMSI_OWNER       = $00000008; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_OWNER}
  PIDMSI_RATING      = $00000009; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_RATING}
  PIDMSI_PRODUCTION  = $0000000A; // VT_FILETIME (UTC)
  {$EXTERNALSYM PIDMSI_PRODUCTION}
  PIDMSI_COPYRIGHT   = $0000000B; // VT_LPWSTR
  {$EXTERNALSYM PIDMSI_COPYRIGHT}


// NtSecApi.h line 566
type
  PLSA_UNICODE_STRING = ^LSA_UNICODE_STRING;
  _LSA_UNICODE_STRING = record
    Length: USHORT;
    MaximumLength: USHORT;
    Buffer: Winapi.Windows.LPWSTR;
  end;
  LSA_UNICODE_STRING = _LSA_UNICODE_STRING;
  TLsaUnicodeString = LSA_UNICODE_STRING;
  PLsaUnicodeString = PLSA_UNICODE_STRING;

  PLSA_STRING = ^LSA_STRING;
  _LSA_STRING = record
    Length: USHORT;
    MaximumLength: USHORT;
    Buffer: PANSICHAR;
  end;
  LSA_STRING = _LSA_STRING;
  TLsaString = LSA_STRING;
  PLsaString = PLSA_STRING;

  PLSA_OBJECT_ATTRIBUTES = ^LSA_OBJECT_ATTRIBUTES;
  _LSA_OBJECT_ATTRIBUTES = record
    Length: ULONG;
    RootDirectory: Winapi.Windows.THandle;
    ObjectName: PLSA_UNICODE_STRING;
    Attributes: ULONG;
    SecurityDescriptor: Pointer; // Points to type SECURITY_DESCRIPTOR
    SecurityQualityOfService: Pointer; // Points to type SECURITY_QUALITY_OF_SERVICE
  end;
  LSA_OBJECT_ATTRIBUTES = _LSA_OBJECT_ATTRIBUTES;
  TLsaObjectAttributes = _LSA_OBJECT_ATTRIBUTES;
  PLsaObjectAttributes = PLSA_OBJECT_ATTRIBUTES;

// NtSecApi.h line 680

////////////////////////////////////////////////////////////////////////////
//                                                                        //
// Local Security Policy Administration API datatypes and defines         //
//                                                                        //
////////////////////////////////////////////////////////////////////////////

//
// Access types for the Policy object
//

const
  POLICY_VIEW_LOCAL_INFORMATION = $00000001;
  {$EXTERNALSYM POLICY_VIEW_LOCAL_INFORMATION}
  POLICY_VIEW_AUDIT_INFORMATION = $00000002;
  {$EXTERNALSYM POLICY_VIEW_AUDIT_INFORMATION}
  POLICY_GET_PRIVATE_INFORMATION = $00000004;
  {$EXTERNALSYM POLICY_GET_PRIVATE_INFORMATION}
  POLICY_TRUST_ADMIN = $00000008;
  {$EXTERNALSYM POLICY_TRUST_ADMIN}
  POLICY_CREATE_ACCOUNT = $00000010;
  {$EXTERNALSYM POLICY_CREATE_ACCOUNT}
  POLICY_CREATE_SECRET = $00000020;
  {$EXTERNALSYM POLICY_CREATE_SECRET}
  POLICY_CREATE_PRIVILEGE = $00000040;
  {$EXTERNALSYM POLICY_CREATE_PRIVILEGE}
  POLICY_SET_DEFAULT_QUOTA_LIMITS = $00000080;
  {$EXTERNALSYM POLICY_SET_DEFAULT_QUOTA_LIMITS}
  POLICY_SET_AUDIT_REQUIREMENTS = $00000100;
  {$EXTERNALSYM POLICY_SET_AUDIT_REQUIREMENTS}
  POLICY_AUDIT_LOG_ADMIN = $00000200;
  {$EXTERNALSYM POLICY_AUDIT_LOG_ADMIN}
  POLICY_SERVER_ADMIN = $00000400;
  {$EXTERNALSYM POLICY_SERVER_ADMIN}
  POLICY_LOOKUP_NAMES = $00000800;
  {$EXTERNALSYM POLICY_LOOKUP_NAMES}
  POLICY_NOTIFICATION = $00001000;
  {$EXTERNALSYM POLICY_NOTIFICATION}

  POLICY_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED or
                               POLICY_VIEW_LOCAL_INFORMATION or
                               POLICY_VIEW_AUDIT_INFORMATION or
                               POLICY_GET_PRIVATE_INFORMATION or
                               POLICY_TRUST_ADMIN or
                               POLICY_CREATE_ACCOUNT or
                               POLICY_CREATE_SECRET or
                               POLICY_CREATE_PRIVILEGE or
                               POLICY_SET_DEFAULT_QUOTA_LIMITS or
                               POLICY_SET_AUDIT_REQUIREMENTS or
                               POLICY_AUDIT_LOG_ADMIN or
                               POLICY_SERVER_ADMIN or
                               POLICY_LOOKUP_NAMES);
  {$EXTERNALSYM POLICY_ALL_ACCESS}

  POLICY_READ = (STANDARD_RIGHTS_READ or
                               POLICY_VIEW_AUDIT_INFORMATION or
                               POLICY_GET_PRIVATE_INFORMATION);
  {$EXTERNALSYM POLICY_READ}

  POLICY_WRITE = (STANDARD_RIGHTS_WRITE or
                               POLICY_TRUST_ADMIN or
                               POLICY_CREATE_ACCOUNT or
                               POLICY_CREATE_SECRET or
                               POLICY_CREATE_PRIVILEGE or
                               POLICY_SET_DEFAULT_QUOTA_LIMITS or
                               POLICY_SET_AUDIT_REQUIREMENTS or
                               POLICY_AUDIT_LOG_ADMIN or
                               POLICY_SERVER_ADMIN);
  {$EXTERNALSYM POLICY_WRITE}

  POLICY_EXECUTE = (STANDARD_RIGHTS_EXECUTE or
                               POLICY_VIEW_LOCAL_INFORMATION or
                               POLICY_LOOKUP_NAMES);
  {$EXTERNALSYM POLICY_EXECUTE}

// NtSecApi.h line 914
type
  _POLICY_INFORMATION_CLASS = (
    picFill0,
    PolicyAuditLogInformation,
    PolicyAuditEventsInformation,
    PolicyPrimaryDomainInformation,
    PolicyPdAccountInformation,
    PolicyAccountDomainInformation,
    PolicyLsaServerRoleInformation,
    PolicyReplicaSourceInformation,
    PolicyDefaultQuotaInformation,
    PolicyModificationInformation,
    PolicyAuditFullSetInformation,
    PolicyAuditFullQueryInformation,
    PolicyDnsDomainInformation,
    PolicyDnsDomainInformationInt);
  {$EXTERNALSYM _POLICY_INFORMATION_CLASS}
  POLICY_INFORMATION_CLASS = _POLICY_INFORMATION_CLASS;
  {$EXTERNALSYM POLICY_INFORMATION_CLASS}
  PPOLICY_INFORMATION_CLASS = ^POLICY_INFORMATION_CLASS;
  {$EXTERNALSYM PPOLICY_INFORMATION_CLASS}
  TPolicyInformationClass = POLICY_INFORMATION_CLASS;
  {$EXTERNALSYM TPolicyInformationClass}
  PPolicyInformationClass = PPOLICY_INFORMATION_CLASS;
  {$EXTERNALSYM PPolicyInformationClass}

// NtSecApi.h line 1031
//
// The following structure corresponds to the PolicyAccountDomainInformation
// information class.
//
type
  PPOLICY_ACCOUNT_DOMAIN_INFO = ^POLICY_ACCOUNT_DOMAIN_INFO;
  _POLICY_ACCOUNT_DOMAIN_INFO = record
    DomainName: LSA_UNICODE_STRING;
    DomainSid: Winapi.Windows.PSID;
  end;
  POLICY_ACCOUNT_DOMAIN_INFO = _POLICY_ACCOUNT_DOMAIN_INFO;
  TPolicyAccountDomainInfo = POLICY_ACCOUNT_DOMAIN_INFO;
  PPolicyAccountDomainInfo = PPOLICY_ACCOUNT_DOMAIN_INFO;

// NtSecApi.h line 1298
type
  LSA_HANDLE = Pointer;
  PLSA_HANDLE = ^LSA_HANDLE;
  TLsaHandle = LSA_HANDLE;

// NtSecApi.h line 1714
type
  NTSTATUS = DWORD;


//
// The th32ProcessID argument is only used if TH32CS_SNAPHEAPLIST or
// TH32CS_SNAPMODULE is specified. th32ProcessID == 0 means the current
// process.
//
// NOTE that all of the snapshots are global except for the heap and module
//      lists which are process specific. To enumerate the heap or module
//      state for all WIN32 processes call with TH32CS_SNAPALL and the
//      current process. Then for each process in the TH32CS_SNAPPROCESS
//      list that isn't the current process, do a call with just
//      TH32CS_SNAPHEAPLIST and/or TH32CS_SNAPMODULE.
//
// dwFlags
//

const
  TH32CS_SNAPHEAPLIST = $00000001;
  {$EXTERNALSYM TH32CS_SNAPHEAPLIST}
  TH32CS_SNAPPROCESS  = $00000002;
  {$EXTERNALSYM TH32CS_SNAPPROCESS}
  TH32CS_SNAPTHREAD   = $00000004;
  {$EXTERNALSYM TH32CS_SNAPTHREAD}
  TH32CS_SNAPMODULE   = $00000008;
  {$EXTERNALSYM TH32CS_SNAPMODULE}
  TH32CS_SNAPMODULE32 = $00000010;
  {$EXTERNALSYM TH32CS_SNAPMODULE32}
  TH32CS_SNAPALL      = TH32CS_SNAPHEAPLIST or TH32CS_SNAPPROCESS or
                        TH32CS_SNAPTHREAD or TH32CS_SNAPMODULE;
  {$EXTERNALSYM TH32CS_SNAPALL}
  TH32CS_INHERIT      = $80000000;
  {$EXTERNALSYM TH32CS_INHERIT}

//
// Use CloseHandle to destroy the snapshot
//

// Thread walking

type
  PTHREADENTRY32 = ^THREADENTRY32;
  {$EXTERNALSYM PTHREADENTRY32}
  tagTHREADENTRY32 = record
    dwSize: DWORD;
    cntUsage: DWORD;
    th32ThreadID: DWORD;       // this thread
    th32OwnerProcessID: DWORD; // Process this thread is associated with
    tpBasePri: Longint;
    tpDeltaPri: Longint;
    dwFlags: DWORD;
  end;
  {$EXTERNALSYM tagTHREADENTRY32}
  THREADENTRY32 = tagTHREADENTRY32;
  {$EXTERNALSYM THREADENTRY32}
  LPTHREADENTRY32 = ^THREADENTRY32;
  {$EXTERNALSYM LPTHREADENTRY32}
  TThreadEntry32 = THREADENTRY32;
  {$EXTERNALSYM TThreadEntry32}



type
  _THREAD_INFORMATION_CLASS = type Cardinal;
  {$EXTERNALSYM _THREAD_INFORMATION_CLASS}
  THREAD_INFORMATION_CLASS = _THREAD_INFORMATION_CLASS;
  {$EXTERNALSYM THREAD_INFORMATION_CLASS}
  PTHREAD_INFORMATION_CLASS = ^_THREAD_INFORMATION_CLASS;
  {$EXTERNALSYM PTHREAD_INFORMATION_CLASS}

const
  ThreadBasicInformation          = 0;
  {$EXTERNALSYM ThreadBasicInformation}
  ThreadTimes                     = 1;
  {$EXTERNALSYM ThreadTimes}
  ThreadPriority                  = 2;
  {$EXTERNALSYM ThreadPriority}
  ThreadBasePriority              = 3;
  {$EXTERNALSYM ThreadBasePriority}
  ThreadAffinityMask              = 4;
  {$EXTERNALSYM ThreadAffinityMask}
  ThreadImpersonationToken        = 5;
  {$EXTERNALSYM ThreadImpersonationToken}
  ThreadDescriptorTableEntry      = 6;
  {$EXTERNALSYM ThreadDescriptorTableEntry}
  ThreadEnableAlignmentFaultFixup = 7;
  {$EXTERNALSYM ThreadEnableAlignmentFaultFixup}
  ThreadEventPair                 = 8;
  {$EXTERNALSYM ThreadEventPair}
  ThreadQuerySetWin32StartAddress = 9;
  {$EXTERNALSYM ThreadQuerySetWin32StartAddress}
  ThreadZeroTlsCell               = 10;
  {$EXTERNALSYM ThreadZeroTlsCell}
  ThreadPerformanceCount          = 11;
  {$EXTERNALSYM ThreadPerformanceCount}
  ThreadAmILastThread             = 12;
  {$EXTERNALSYM ThreadAmILastThread}
  ThreadIdealProcessor            = 13;
  {$EXTERNALSYM ThreadIdealProcessor}
  ThreadPriorityBoost             = 14;
  {$EXTERNALSYM ThreadPriorityBoost}
  ThreadSetTlsArrayAddress        = 15;
  {$EXTERNALSYM ThreadSetTlsArrayAddress}
  ThreadIsIoPending               = 16;
  {$EXTERNALSYM ThreadIsIoPending}
  ThreadHideFromDebugger          = 17;
  {$EXTERNALSYM ThreadHideFromDebugger}

type
  _CLIENT_ID = record
    UniqueProcess: THandle;
    UniqueThread: THandle;
  end;
  {$EXTERNALSYM _CLIENT_ID}
  CLIENT_ID = _CLIENT_ID;
  {$EXTERNALSYM CLIENT_ID}
  PCLIENT_ID = ^CLIENT_ID;
  {$EXTERNALSYM PCLIENT_ID}

  KAFFINITY = ULONG;
  {$EXTERNALSYM KAFFINITY}

  KPRIORITY = LongInt;
  {$EXTERNALSYM KPRIORITY}

  _THREAD_BASIC_INFORMATION = record
    ExitStatus: NTSTATUS;
    TebBaseAddress: Pointer;
    ClientId: CLIENT_ID;
    AffinityMask: KAFFINITY;
    Priority: KPRIORITY;
    BasePriority: KPRIORITY;
  end;
  {$EXTERNALSYM _THREAD_BASIC_INFORMATION}
  THREAD_BASIC_INFORMATION = _THREAD_BASIC_INFORMATION;
  {$EXTERNALSYM THREAD_BASIC_INFORMATION}
  PTHREAD_BASIC_INFORMATION = ^_THREAD_BASIC_INFORMATION;
  {$EXTERNALSYM PTHREAD_BASIC_INFORMATION}



//DOM-IGNORE-END


const
  RtdlSetWaitableTimer: function(hTimer: THandle; var lpDueTime: TLargeInteger;
    lPeriod: Longint; pfnCompletionRoutine: TFNTimerAPCRoutine;
    lpArgToCompletionRoutine: Pointer; fResume: BOOL): BOOL stdcall = SetWaitableTimer;


function NtQueryInformationThread(ThreadHandle: THandle; ThreadInformationClass: THREAD_INFORMATION_CLASS;
  ThreadInformation: Pointer; ThreadInformationLength: ULONG; ReturnLength: PULONG): NTSTATUS; stdcall;
{$EXTERNALSYM NtQueryInformationThread}

function CaptureStackBackTrace(FramesToSkip, FramesToCapture: DWORD;
  BackTrace: Pointer; out BackTraceHash: DWORD): Word; stdcall;
{$EXTERNALSYM CaptureStackBackTrace}

type
  // Smart name compare function
  TJclSmartCompOption = (scSimpleCompare, scIgnoreCase);
  TJclSmartCompOptions = set of TJclSmartCompOption;

function PeStripFunctionAW(const FunctionName: string): string;

function PeSmartFunctionNameSame(const ComparedName, FunctionName: string;
  Options: TJclSmartCompOptions = []): Boolean;

const
  IMAGE_FILE_SYSTEM        = $1000; // System File.
  IMAGE_FILE_MACHINE_I386  = $014c; // Intel 386.
  IMAGE_FILE_MACHINE_AMD64 = $8664; // AMD64 (K8)
const
  IMAGE_ORDINAL_FLAG64 = ULONGLONG($8000000000000000);
  IMAGE_ORDINAL_FLAG32 = DWORD($80000000);
//
// Based relocation format.
//

type
  PIMAGE_BASE_RELOCATION = ^IMAGE_BASE_RELOCATION;
  {$EXTERNALSYM PIMAGE_BASE_RELOCATION}
  _IMAGE_BASE_RELOCATION = record
    VirtualAddress: DWORD;
    SizeOfBlock: DWORD;
    //  WORD    TypeOffset[1];
  end;
  {$EXTERNALSYM _IMAGE_BASE_RELOCATION}
  IMAGE_BASE_RELOCATION = _IMAGE_BASE_RELOCATION;
  {$EXTERNALSYM IMAGE_BASE_RELOCATION}
  TImageBaseRelocation = IMAGE_BASE_RELOCATION;
  PImageBaseRelocation = PIMAGE_BASE_RELOCATION;

  PIMAGE_EXPORT_DIRECTORY = ^IMAGE_EXPORT_DIRECTORY;
  {$EXTERNALSYM PIMAGE_EXPORT_DIRECTORY}
  _IMAGE_EXPORT_DIRECTORY = record
    Characteristics: DWORD;
    TimeDateStamp: DWORD;
    MajorVersion: Word;
    MinorVersion: Word;
    Name: DWORD;
    Base: DWORD;
    NumberOfFunctions: DWORD;
    NumberOfNames: DWORD;
    AddressOfFunctions: DWORD; // RVA from base of image
    AddressOfNames: DWORD; // RVA from base of image
    AddressOfNameOrdinals: DWORD; // RVA from base of image
  end;
  {$EXTERNALSYM _IMAGE_EXPORT_DIRECTORY}
  IMAGE_EXPORT_DIRECTORY = _IMAGE_EXPORT_DIRECTORY;
  {$EXTERNALSYM IMAGE_EXPORT_DIRECTORY}
  TImageExportDirectory = IMAGE_EXPORT_DIRECTORY;
  PImageExportDirectory = PIMAGE_EXPORT_DIRECTORY;

//
// Import Format
//
  PIMAGE_IMPORT_BY_NAME = ^IMAGE_IMPORT_BY_NAME;
  {$EXTERNALSYM PIMAGE_IMPORT_BY_NAME}
  _IMAGE_IMPORT_BY_NAME = record
    Hint: Word;
    Name: array [0..0] of AnsiChar;
  end;
  {$EXTERNALSYM _IMAGE_IMPORT_BY_NAME}
  IMAGE_IMPORT_BY_NAME = _IMAGE_IMPORT_BY_NAME;
  {$EXTERNALSYM IMAGE_IMPORT_BY_NAME}
  TImageImportByName = IMAGE_IMPORT_BY_NAME;
  PImageImportByName = PIMAGE_IMPORT_BY_NAME;

  PIMAGE_THUNK_DATA64 = ^IMAGE_THUNK_DATA64;
  {$EXTERNALSYM PIMAGE_THUNK_DATA64}
  _IMAGE_THUNK_DATA64 = record
    case Integer of
      0: (ForwarderString: ULONGLONG);   // PBYTE
      1: (Function_: ULONGLONG);         // PDWORD
      2: (Ordinal: ULONGLONG);
      3: (AddressOfData: ULONGLONG);     // PIMAGE_IMPORT_BY_NAME
  end;
  {$EXTERNALSYM _IMAGE_THUNK_DATA64}
  IMAGE_THUNK_DATA64 = _IMAGE_THUNK_DATA64;
  {$EXTERNALSYM IMAGE_THUNK_DATA64}
  TImageThunkData64 = IMAGE_THUNK_DATA64;
  PImageThunkData64 = PIMAGE_THUNK_DATA64;

// #include "poppack.h"                        // Back to 4 byte packing

  PIMAGE_THUNK_DATA32 = ^IMAGE_THUNK_DATA32;
  {$EXTERNALSYM PIMAGE_THUNK_DATA32}
  _IMAGE_THUNK_DATA32 = record
    case Integer of
      0: (ForwarderString: DWORD);   // PBYTE
      1: (Function_: DWORD);         // PDWORD
      2: (Ordinal: DWORD);
      3: (AddressOfData: DWORD);     // PIMAGE_IMPORT_BY_NAME
  end;
  {$EXTERNALSYM _IMAGE_THUNK_DATA32}
  IMAGE_THUNK_DATA32 = _IMAGE_THUNK_DATA32;
  {$EXTERNALSYM IMAGE_THUNK_DATA32}
  TImageThunkData32 = IMAGE_THUNK_DATA32;
  PImageThunkData32 = PIMAGE_THUNK_DATA32;

type
  TIIDUnion = record
    case Integer of
      0: (Characteristics: DWORD);         // 0 for terminating null import descriptor
      1: (OriginalFirstThunk: DWORD);      // RVA to original unbound IAT (PIMAGE_THUNK_DATA)
  end;

  PIMAGE_IMPORT_DESCRIPTOR = ^IMAGE_IMPORT_DESCRIPTOR;
  {$EXTERNALSYM PIMAGE_IMPORT_DESCRIPTOR}
  _IMAGE_IMPORT_DESCRIPTOR = record
    Union: TIIDUnion;
    TimeDateStamp: DWORD;                  // 0 if not bound,
                                           // -1 if bound, and real date\time stamp
                                           //     in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                                           // O.W. date/time stamp of DLL bound to (Old BIND)

    ForwarderChain: DWORD;                 // -1 if no forwarders
    Name: DWORD;
    FirstThunk: DWORD;                     // RVA to IAT (if bound this IAT has actual addresses)
  end;
  {$EXTERNALSYM _IMAGE_IMPORT_DESCRIPTOR}
  IMAGE_IMPORT_DESCRIPTOR = _IMAGE_IMPORT_DESCRIPTOR;
  {$EXTERNALSYM IMAGE_IMPORT_DESCRIPTOR}
  TImageImportDescriptor = IMAGE_IMPORT_DESCRIPTOR;
  PImageImportDescriptor = PIMAGE_IMPORT_DESCRIPTOR;

type
  PIMAGE_BOUND_IMPORT_DESCRIPTOR = ^IMAGE_BOUND_IMPORT_DESCRIPTOR;
  {$EXTERNALSYM PIMAGE_BOUND_IMPORT_DESCRIPTOR}
  _IMAGE_BOUND_IMPORT_DESCRIPTOR = record
    TimeDateStamp: DWORD;
    OffsetModuleName: Word;
    NumberOfModuleForwarderRefs: Word;
    // Array of zero or more IMAGE_BOUND_FORWARDER_REF follows
  end;
  {$EXTERNALSYM _IMAGE_BOUND_IMPORT_DESCRIPTOR}
  IMAGE_BOUND_IMPORT_DESCRIPTOR = _IMAGE_BOUND_IMPORT_DESCRIPTOR;
  {$EXTERNALSYM IMAGE_BOUND_IMPORT_DESCRIPTOR}
  TImageBoundImportDescriptor = IMAGE_BOUND_IMPORT_DESCRIPTOR;
  PImageBoundImportDescriptor = PIMAGE_BOUND_IMPORT_DESCRIPTOR;
type

  PIMAGE_BOUND_FORWARDER_REF = ^IMAGE_BOUND_FORWARDER_REF;
  {$EXTERNALSYM PIMAGE_BOUND_FORWARDER_REF}
  _IMAGE_BOUND_FORWARDER_REF = record
    TimeDateStamp: DWORD;
    OffsetModuleName: Word;
    Reserved: Word;
  end;
  {$EXTERNALSYM _IMAGE_BOUND_FORWARDER_REF}
  IMAGE_BOUND_FORWARDER_REF = _IMAGE_BOUND_FORWARDER_REF;
  {$EXTERNALSYM IMAGE_BOUND_FORWARDER_REF}
  TImageBoundForwarderRef = IMAGE_BOUND_FORWARDER_REF;
  PImageBoundForwarderRef = PIMAGE_BOUND_FORWARDER_REF;

  PIMAGE_RESOURCE_DIRECTORY = ^IMAGE_RESOURCE_DIRECTORY;
  {$EXTERNALSYM PIMAGE_RESOURCE_DIRECTORY}
  _IMAGE_RESOURCE_DIRECTORY = record
    Characteristics: DWORD;
    TimeDateStamp: DWORD;
    MajorVersion: Word;
    MinorVersion: Word;
    NumberOfNamedEntries: Word;
    NumberOfIdEntries: Word;
    // IMAGE_RESOURCE_DIRECTORY_ENTRY DirectoryEntries[];
  end;
  {$EXTERNALSYM _IMAGE_RESOURCE_DIRECTORY}
  IMAGE_RESOURCE_DIRECTORY = _IMAGE_RESOURCE_DIRECTORY;
  {$EXTERNALSYM IMAGE_RESOURCE_DIRECTORY}
  TImageResourceDirectory = IMAGE_RESOURCE_DIRECTORY;
  PImageResourceDirectory = PIMAGE_RESOURCE_DIRECTORY;

const
  IMAGE_RESOURCE_NAME_IS_STRING    = DWORD($80000000);
  {$EXTERNALSYM IMAGE_RESOURCE_NAME_IS_STRING}
  IMAGE_RESOURCE_DATA_IS_DIRECTORY = DWORD($80000000);
  {$EXTERNALSYM IMAGE_RESOURCE_DATA_IS_DIRECTORY}

type
  PIMAGE_RESOURCE_DIRECTORY_ENTRY = ^IMAGE_RESOURCE_DIRECTORY_ENTRY;
  {$EXTERNALSYM PIMAGE_RESOURCE_DIRECTORY_ENTRY}
  _IMAGE_RESOURCE_DIRECTORY_ENTRY = record
    case Integer of
      0: (
        // DWORD NameOffset:31;
        // DWORD NameIsString:1;
        NameOffset: DWORD;
        OffsetToData: DWORD
      );
      1: (
        Name: DWORD;
        // DWORD OffsetToDirectory:31;
        // DWORD DataIsDirectory:1;
        OffsetToDirectory: DWORD;
      );
      2: (
        Id: WORD;
      );
  end;
  {$EXTERNALSYM _IMAGE_RESOURCE_DIRECTORY_ENTRY}
  IMAGE_RESOURCE_DIRECTORY_ENTRY = _IMAGE_RESOURCE_DIRECTORY_ENTRY;
  {$EXTERNALSYM IMAGE_RESOURCE_DIRECTORY_ENTRY}
  TImageResourceDirectoryEntry = IMAGE_RESOURCE_DIRECTORY_ENTRY;
  PImageResourceDirectoryEntry = PIMAGE_RESOURCE_DIRECTORY_ENTRY;

  PIMAGE_RESOURCE_DIR_STRING_U = ^IMAGE_RESOURCE_DIR_STRING_U;
  {$EXTERNALSYM PIMAGE_RESOURCE_DIR_STRING_U}
  _IMAGE_RESOURCE_DIR_STRING_U = record
    Length: Word;
    NameString: array [0..0] of WCHAR;
  end;
  {$EXTERNALSYM _IMAGE_RESOURCE_DIR_STRING_U}
  IMAGE_RESOURCE_DIR_STRING_U = _IMAGE_RESOURCE_DIR_STRING_U;
  {$EXTERNALSYM IMAGE_RESOURCE_DIR_STRING_U}
  TImageResourceDirStringU = IMAGE_RESOURCE_DIR_STRING_U;
  PImageResourceDirStringU = PIMAGE_RESOURCE_DIR_STRING_U;

  PIMAGE_RESOURCE_DATA_ENTRY = ^IMAGE_RESOURCE_DATA_ENTRY;
  {$EXTERNALSYM PIMAGE_RESOURCE_DATA_ENTRY}
  _IMAGE_RESOURCE_DATA_ENTRY = record
    OffsetToData: DWORD;
    Size: DWORD;
    CodePage: DWORD;
    Reserved: DWORD;
  end;
  {$EXTERNALSYM _IMAGE_RESOURCE_DATA_ENTRY}
  IMAGE_RESOURCE_DATA_ENTRY = _IMAGE_RESOURCE_DATA_ENTRY;
  {$EXTERNALSYM IMAGE_RESOURCE_DATA_ENTRY}
  TImageResourceDataEntry = IMAGE_RESOURCE_DATA_ENTRY;
  PImageResourceDataEntry = PIMAGE_RESOURCE_DATA_ENTRY;

//
// Load Configuration Directory Entry
//

type
  PIMAGE_LOAD_CONFIG_DIRECTORY32 = ^IMAGE_LOAD_CONFIG_DIRECTORY32;
  {$EXTERNALSYM PIMAGE_LOAD_CONFIG_DIRECTORY32}
  IMAGE_LOAD_CONFIG_DIRECTORY32 = record
    Size: DWORD;
    TimeDateStamp: DWORD;
    MajorVersion: WORD;
    MinorVersion: WORD;
    GlobalFlagsClear: DWORD;
    GlobalFlagsSet: DWORD;
    CriticalSectionDefaultTimeout: DWORD;
    DeCommitFreeBlockThreshold: DWORD;
    DeCommitTotalFreeThreshold: DWORD;
    LockPrefixTable: DWORD;            // VA
    MaximumAllocationSize: DWORD;
    VirtualMemoryThreshold: DWORD;
    ProcessHeapFlags: DWORD;
    ProcessAffinityMask: DWORD;
    CSDVersion: WORD;
    Reserved1: WORD;
    EditList: DWORD;                   // VA
    SecurityCookie: DWORD;             // VA
    SEHandlerTable: DWORD;             // VA
    SEHandlerCount: DWORD;
  end;
  {$EXTERNALSYM IMAGE_LOAD_CONFIG_DIRECTORY32}
  TImageLoadConfigDirectory32 = IMAGE_LOAD_CONFIG_DIRECTORY32;
  PImageLoadConfigDirectory32 = PIMAGE_LOAD_CONFIG_DIRECTORY32;

  PIMAGE_LOAD_CONFIG_DIRECTORY64 = ^IMAGE_LOAD_CONFIG_DIRECTORY64;
  {$EXTERNALSYM PIMAGE_LOAD_CONFIG_DIRECTORY64}
  IMAGE_LOAD_CONFIG_DIRECTORY64 = record
    Size: DWORD;
    TimeDateStamp: DWORD;
    MajorVersion: WORD;
    MinorVersion: WORD;
    GlobalFlagsClear: DWORD;
    GlobalFlagsSet: DWORD;
    CriticalSectionDefaultTimeout: DWORD;
    DeCommitFreeBlockThreshold: ULONGLONG;
    DeCommitTotalFreeThreshold: ULONGLONG;
    LockPrefixTable: ULONGLONG;         // VA
    MaximumAllocationSize: ULONGLONG;
    VirtualMemoryThreshold: ULONGLONG;
    ProcessAffinityMask: ULONGLONG;
    ProcessHeapFlags: DWORD;
    CSDVersion: WORD;
    Reserved1: WORD;
    EditList: ULONGLONG;                // VA
    SecurityCookie: ULONGLONG;             // VA
    SEHandlerTable: ULONGLONG;             // VA
    SEHandlerCount: ULONGLONG;
  end;
  {$EXTERNALSYM IMAGE_LOAD_CONFIG_DIRECTORY64}
  TImageLoadConfigDirectory64 = IMAGE_LOAD_CONFIG_DIRECTORY64;
  PImageLoadConfigDirectory64 = PIMAGE_LOAD_CONFIG_DIRECTORY64;

  IMAGE_LOAD_CONFIG_DIRECTORY = IMAGE_LOAD_CONFIG_DIRECTORY32;
  {$EXTERNALSYM IMAGE_LOAD_CONFIG_DIRECTORY}
  PIMAGE_LOAD_CONFIG_DIRECTORY = PIMAGE_LOAD_CONFIG_DIRECTORY32;
  {$EXTERNALSYM PIMAGE_LOAD_CONFIG_DIRECTORY}
  TImageLoadConfigDirectory = TImageLoadConfigDirectory32;
  PImageLoadConfigDirectory = PImageLoadConfigDirectory32;

type
  IMAGE_COR20_HEADER = record

    // Header versioning

    cb: DWORD;
    MajorRuntimeVersion: WORD;
    MinorRuntimeVersion: WORD;

    // Symbol table and startup information

    MetaData: IMAGE_DATA_DIRECTORY;
    Flags: DWORD;
    EntryPointToken: DWORD;

    // Binding information

    Resources: IMAGE_DATA_DIRECTORY;
    StrongNameSignature: IMAGE_DATA_DIRECTORY;

    // Regular fixup and binding information

    CodeManagerTable: IMAGE_DATA_DIRECTORY;
    VTableFixups: IMAGE_DATA_DIRECTORY;
    ExportAddressTableJumps: IMAGE_DATA_DIRECTORY;

    // Precompiled image info (internal use only - set to zero)

    ManagedNativeHeader: IMAGE_DATA_DIRECTORY;
  end;
  PIMAGE_COR20_HEADER = ^IMAGE_COR20_HEADER;
  TImageCor20Header = IMAGE_COR20_HEADER;
  PImageCor20Header = PIMAGE_COR20_HEADER;


const
  Borland32BitSymbolFileSignatureForDelphi = $39304246; // 'FB09'
  Borland32BitSymbolFileSignatureForBCB    = $41304246; // 'FB0A'

type
  { Signature structure }
  PJclTD32FileSignature = ^TJclTD32FileSignature;
  TJclTD32FileSignature = packed record
    Signature: DWORD;
    Offset: DWORD;
  end;

const
  { Subsection Types }
  SUBSECTION_TYPE_MODULE         = $120;
  SUBSECTION_TYPE_TYPES          = $121;
  SUBSECTION_TYPE_SYMBOLS        = $124;
  SUBSECTION_TYPE_ALIGN_SYMBOLS  = $125;
  SUBSECTION_TYPE_SOURCE_MODULE  = $127;
  SUBSECTION_TYPE_GLOBAL_SYMBOLS = $129;
  SUBSECTION_TYPE_GLOBAL_TYPES   = $12B;
  SUBSECTION_TYPE_NAMES          = $130;

type
  { Subsection directory header structure }
  { The directory header structure is followed by the directory entries
    which specify the subsection type, module index, file offset, and size.
    The subsection directory gives the location (LFO) and size of each subsection,
    as well as its type and module number if applicable. }
  PDirectoryEntry = ^TDirectoryEntry;
  TDirectoryEntry = packed record
    SubsectionType: Word; // Subdirectory type
    ModuleIndex: Word;    // Module index
    Offset: DWORD;        // Offset from the base offset lfoBase
    Size: DWORD;          // Number of bytes in subsection
  end;

  { The subsection directory is prefixed with a directory header structure
    indicating size and number of subsection directory entries that follow. }
  PDirectoryHeader = ^TDirectoryHeader;
  TDirectoryHeader = packed record
    Size: Word;           // Length of this structure
    DirEntrySize: Word;   // Length of each directory entry
    DirEntryCount: DWORD; // Number of directory entries
    lfoNextDir: DWORD;    // Offset from lfoBase of next directory.
    Flags: DWORD;         // Flags describing directory and subsection tables.
    DirEntries: array [0..0] of TDirectoryEntry;
  end;


{*******************************************************************************

  SUBSECTION_TYPE_MODULE $120

  This describes the basic information about an object module including  code
  segments, module name,  and the  number of  segments for  the modules  that
  follow.  Directory entries for  sstModules  precede  all  other  subsection
  directory entries.

*******************************************************************************}

type
  PSegmentInfo = ^TSegmentInfo;
  TSegmentInfo = packed record
    Segment: Word; // Segment that this structure describes
    Flags: Word;   // Attributes for the logical segment.
                   // The following attributes are defined:
                   //   $0000  Data segment
                   //   $0001  Code segment
    Offset: DWORD; // Offset in segment where the code starts
    Size: DWORD;   // Count of the number of bytes of code in the segment
  end;
  PSegmentInfoArray = ^TSegmentInfoArray;
  TSegmentInfoArray = array [0..32767] of TSegmentInfo;

  PModuleInfo = ^TModuleInfo1;
  TModuleInfo1 = packed record
    OverlayNumber: Word;  // Overlay number
    LibraryIndex: Word;   // Index into sstLibraries subsection
                          // if this module was linked from a library
    SegmentCount: Word;   // Count of the number of code segments
                          // this module contributes to
    DebuggingStyle: Word; // Debugging style  for this  module.
    NameIndex: DWORD;     // Name index of module.
    TimeStamp: DWORD;     // Time stamp from the OBJ file.
    Reserved: array [0..2] of DWORD; // Set to 0.
    Segments: array [0..0] of TSegmentInfo;
                          // Detailed information about each segment
                          // that code is contributed to.
                          // This is an array of cSeg count segment
                          // information descriptor structures.
  end;

{*******************************************************************************

  SUBSECTION_TYPE_SOURCE_MODULE $0127

  This table describes the source line number to addressing mapping
  information for a module. The table permits the description of a module
  containing multiple source files with each source file contributing code to
  one or more code segments. The base addresses of the tables described
  below are all relative to the beginning of the sstSrcModule table.


  Module header

  Information for source file 1

    Information for segment 1
         .
         .
         .
    Information for segment n

         .
         .
         .

  Information for source file n

    Information for segment 1
         .
         .
         .
    Information for segment n

*******************************************************************************}
type
  { The line number to address mapping information is contained in a table with
    the following format: }
  PLineMappingEntry = ^TLineMappingEntry;
  TLineMappingEntry = packed record
    SegmentIndex: Word;  // Segment index for this table
    PairCount: Word;     // Count of the number of source line pairs to follow
    Offsets: array [0..0] of DWORD;
                     // An array of 32-bit offsets for the offset
                     // within the code segment ofthe start of ine contained
                     // in the parallel array linenumber.
    (*
    { This is an array of 16-bit line numbers of the lines in the source file
      that cause code to be emitted to the code segment.
      This array is parallel to the offset array.
      If cPair is not even, then a zero word is emitted to
      maintain natural alignment in the sstSrcModule table. }
    LineNumbers: array [0..PairCount - 1] of Word;
    *)
  end;

  TOffsetPair = packed record
    StartOffset: DWORD;
    EndOffset: DWORD;
  end;
  POffsetPairArray = ^TOffsetPairArray;
  TOffsetPairArray = array [0..32767] of TOffsetPair;

  { The file table describes the code segments that receive code from this
    source file. Source file entries have the following format: }
  PSourceFileEntry = ^TSourceFileEntry;
  TSourceFileEntry = packed record
    SegmentCount: Word; // Number of segments that receive code from this source file.
    NameIndex: DWORD;   // Name index of Source file name.

    BaseSrcLines: array [0..0] of DWORD;
                        // An array of offsets for the line/address mapping
                        // tables for each of the segments that receive code
                        // from this source file.
    (*
    { An array  of two  32-bit offsets  per segment  that
      receives code from this  module.  The first  offset
      is the offset within the segment of the first  byte
      of code from this module.  The second offset is the
      ending address of the  code from this module.   The
      order of these pairs corresponds to the ordering of
      the segments in the  seg  array.   Zeros  in  these
      entries means that the information is not known and
      the file and line tables described below need to be
      examined to determine if an address of interest  is
      contained within the code from this module. }
    SegmentAddress: array [0..SegmentCount - 1] of TOffsetPair;

    Name: ShortString; // Count of the number of bytes in source file name
    *)
  end;

  { The module header structure describes the source file and code segment
    organization of the module. Each module header has the following format: }
  PSourceModuleInfo = ^TSourceModuleInfo;
  TSourceModuleInfo = packed record
    FileCount: Word;    // The number of source file scontributing code to segments
    SegmentCount: Word; // The number of code segments receiving code from this module

    BaseSrcFiles: array [0..0] of DWORD;
    (*
    // This is an array of base offsets from the beginning of the sstSrcModule table
    BaseSrcFiles: array [0..FileCount - 1] of DWORD;

    { An array  of two  32-bit offsets  per segment  that
      receives code from this  module.  The first  offset
      is the offset within the segment of the first  byte
      of code from this module.  The second offset is the
      ending address of the  code from this module.   The
      order of these pairs corresponds to the ordering of
      the segments in the  seg  array.   Zeros  in  these
      entries means that the information is not known and
      the file and line tables described below need to be
      examined to determine if an address of interest  is
      contained within the code from this module. }
    SegmentAddress: array [0..SegmentCount - 1] of TOffsetPair;

    { An array of segment indices that receive code  from
      this module.  If the  number  of  segments  is  not
      even, a pad word  is inserted  to maintain  natural
      alignment. }
    SegmentIndexes: array [0..SegmentCount - 1] of Word;
    *)
  end;

{*******************************************************************************

  SUBSECTION_TYPE_GLOBAL_TYPES $12b

  This subsection contains the packed  type records for the executable  file.
  The first long word of the subsection  contains the number of types in  the
  table.  This count is  followed by a count-sized  array of long offsets  to
  the  corresponding  type  record.   As  the  sstGlobalTypes  subsection  is
  written, each  type record  is forced  to start  on a  long word  boundary.
  However, the length of the  type string is NOT  adjusted by the pad  count.
  The remainder of the subsection contains  the type records.

*******************************************************************************}

type
  PGlobalTypeInfo = ^TGlobalTypeInfo;
  TGlobalTypeInfo = packed record
    Count: DWORD; // count of the number of types
    // offset of each type string from the beginning of table
    Offsets: array [0..0] of DWORD;
  end;

const
  { Symbol type defines }
  SYMBOL_TYPE_COMPILE        = $0001; // Compile flags symbol
  SYMBOL_TYPE_REGISTER       = $0002; // Register variable
  SYMBOL_TYPE_CONST          = $0003; // Constant symbol
  SYMBOL_TYPE_UDT            = $0004; // User-defined Type
  SYMBOL_TYPE_SSEARCH        = $0005; // Start search
  SYMBOL_TYPE_END            = $0006; // End block, procedure, with, or thunk
  SYMBOL_TYPE_SKIP           = $0007; // Skip - Reserve symbol space
  SYMBOL_TYPE_CVRESERVE      = $0008; // Reserved for Code View internal use
  SYMBOL_TYPE_OBJNAME        = $0009; // Specify name of object file

  SYMBOL_TYPE_BPREL16        = $0100; // BP relative 16:16
  SYMBOL_TYPE_LDATA16        = $0101; // Local data 16:16
  SYMBOL_TYPE_GDATA16        = $0102; // Global data 16:16
  SYMBOL_TYPE_PUB16          = $0103; // Public symbol 16:16
  SYMBOL_TYPE_LPROC16        = $0104; // Local procedure start 16:16
  SYMBOL_TYPE_GPROC16        = $0105; // Global procedure start 16:16
  SYMBOL_TYPE_THUNK16        = $0106; // Thunk start 16:16
  SYMBOL_TYPE_BLOCK16        = $0107; // Block start 16:16
  SYMBOL_TYPE_WITH16         = $0108; // With start 16:16
  SYMBOL_TYPE_LABEL16        = $0109; // Code label 16:16
  SYMBOL_TYPE_CEXMODEL16     = $010A; // Change execution model 16:16
  SYMBOL_TYPE_VFTPATH16      = $010B; // Virtual function table path descriptor 16:16

  SYMBOL_TYPE_BPREL32        = $0200; // BP relative 16:32
  SYMBOL_TYPE_LDATA32        = $0201; // Local data 16:32
  SYMBOL_TYPE_GDATA32        = $0202; // Global data 16:32
  SYMBOL_TYPE_PUB32          = $0203; // Public symbol 16:32
  SYMBOL_TYPE_LPROC32        = $0204; // Local procedure start 16:32
  SYMBOL_TYPE_GPROC32        = $0205; // Global procedure start 16:32
  SYMBOL_TYPE_THUNK32        = $0206; // Thunk start 16:32
  SYMBOL_TYPE_BLOCK32        = $0207; // Block start 16:32
  SYMBOL_TYPE_WITH32         = $0208; // With start 16:32
  SYMBOL_TYPE_LABEL32        = $0209; // Label 16:32
  SYMBOL_TYPE_CEXMODEL32     = $020A; // Change execution model 16:32
  SYMBOL_TYPE_VFTPATH32      = $020B; // Virtual function table path descriptor 16:32

{*******************************************************************************

  Global and Local Procedure Start 16:32

  SYMBOL_TYPE_LPROC32 $0204
  SYMBOL_TYPE_GPROC32 $0205

    The symbol records define local (file static) and global procedure
  definition. For C/C++, functions that are declared static to a module are
  emitted as Local Procedure symbols. Functions not specifically declared
  static are emitted as Global Procedures.
    For each SYMBOL_TYPE_GPROC32 emitted, an SYMBOL_TYPE_GPROCREF symbol
  must be fabricated and emitted to the SUBSECTION_TYPE_GLOBAL_SYMBOLS section.

*******************************************************************************}

type
  TSymbolProcInfo = packed record
    pParent: DWORD;
    pEnd: DWORD;
    pNext: DWORD;
    Size: DWORD;        // Length in bytes of this procedure
    DebugStart: DWORD;  // Offset in bytes from the start of the procedure to
                        // the point where the stack frame has been set up.
    DebugEnd: DWORD;    // Offset in bytes from the start of the procedure to
                        // the point where the  procedure is  ready to  return
                        // and has calculated its return value, if any.
                        // Frame and register variables an still be viewed.
    Offset: DWORD;      // Offset portion of  the segmented address of
                        // the start of the procedure in the code segment
    Segment: Word;      // Segment portion of the segmented address of
                        // the start of the procedure in the code segment
    ProcType: DWORD;    // Type of the procedure type record
    NearFar: Byte;      // Type of return the procedure makes:
                        //   0       near
                        //   4       far
    Reserved: Byte;
    NameIndex: DWORD;   // Name index of procedure
  end;

  TSymbolObjNameInfo = packed record
    Signature: DWORD;   // Signature for the CodeView information contained in
                        // this module
    NameIndex: DWORD;   // Name index of the object file
  end;

  TSymbolDataInfo = packed record
    Offset: DWORD;      // Offset portion of  the segmented address of
                        // the start of the data in the code segment
    Segment: Word;      // Segment portion of the segmented address of
                        // the start of the data in the code segment
    Reserved: Word;
    TypeIndex: DWORD;   // Type index of the symbol
    NameIndex: DWORD;   // Name index of the symbol
  end;

  TSymbolWithInfo = packed record
    pParent: DWORD;
    pEnd: DWORD;
    Size: DWORD;        // Length in bytes of this "with"
    Offset: DWORD;      // Offset portion of the segmented address of
                        // the start of the "with" in the code segment
    Segment: Word;      // Segment portion of the segmented address of
                        // the start of the "with" in the code segment
    Reserved: Word;
    NameIndex: DWORD;   // Name index of the "with"
  end;

  TSymbolLabelInfo = packed record
    Offset: DWORD;      // Offset portion of  the segmented address of
                        // the start of the label in the code segment
    Segment: Word;      // Segment portion of the segmented address of
                        // the start of the label in the code segment
    NearFar: Byte;      // Address mode of the label:
                        //   0       near
                        //   4       far
    Reserved: Byte;
    NameIndex: DWORD;   // Name index of the label
  end;

  TSymbolConstantInfo = packed record
    TypeIndex: DWORD;   // Type index of the constant (for enums)
    NameIndex: DWORD;   // Name index of the constant
    Reserved: DWORD;
    Value: DWORD;       // value of the constant
  end;

  TSymbolUdtInfo = packed record
    TypeIndex: DWORD;   // Type index of the type
    Properties: Word;   // isTag:1 True if this is a tag (not a typedef)
                        // isNest:1 True if the type is a nested type (its name
                        // will be 'class_name::type_name' in that case)
    NameIndex: DWORD;   // Name index of the type
    Reserved: DWORD;
  end;

  TSymbolVftPathInfo = packed record
    Offset: DWORD;      // Offset portion of start of the virtual function table
    Segment: Word;      // Segment portion of the virtual function table
    Reserved: Word;
    RootIndex: DWORD;   // The type index of the class at the root of the path
    PathIndex: DWORD;   // Type index of the record describing the base class
                        // path from the root to the leaf class for the virtual
                        // function table
  end;

type
  { Symbol Information Records }
  PSymbolInfo = ^TSymbolInfo;
  TSymbolInfo = packed record
    Size: Word;
    SymbolType: Word;
    case Word of
      SYMBOL_TYPE_LPROC32, SYMBOL_TYPE_GPROC32:
        (Proc: TSymbolProcInfo);
      SYMBOL_TYPE_OBJNAME:
        (ObjName: TSymbolObjNameInfo);
      SYMBOL_TYPE_LDATA32, SYMBOL_TYPE_GDATA32, SYMBOL_TYPE_PUB32:
        (Data: TSymbolDataInfo);
      SYMBOL_TYPE_WITH32:
        (With32: TSymbolWithInfo);
      SYMBOL_TYPE_LABEL32:
        (Label32: TSymbolLabelInfo);
      SYMBOL_TYPE_CONST:
        (Constant: TSymbolConstantInfo);
      SYMBOL_TYPE_UDT:
        (Udt: TSymbolUdtInfo);
      SYMBOL_TYPE_VFTPATH32:
        (VftPath: TSymbolVftPathInfo);
  end;

  PSymbolInfos = ^TSymbolInfos;
  TSymbolInfos = packed record
    Signature: DWORD;
    Symbols: array [0..0] of TSymbolInfo;
  end;

// TD32 information related classes
type
  TJclTD32ModuleInfo = class(TObject)
  private
    FNameIndex: DWORD;
    FSegments: PSegmentInfoArray;
    FSegmentCount: Integer;
    function GetSegment(const Idx: Integer): TSegmentInfo;
  public
    constructor Create(pModInfo: PModuleInfo);
    property NameIndex: DWORD read FNameIndex;
    property SegmentCount: Integer read FSegmentCount; //GetSegmentCount;
    property Segment[const Idx: Integer]: TSegmentInfo read GetSegment; default;
  end;

  TJclTD32LineInfo = class(TObject)
  private
    FLineNo: DWORD;
    FOffset: DWORD;
  public
    constructor Create(ALineNo, AOffset: DWORD);
    property LineNo: DWORD read FLineNo;
    property Offset: DWORD read FOffset;
  end;

  TJclTD32SourceModuleInfo = class(TObject)
  private
    FLines: TObjectList;
    FSegments: POffsetPairArray;
    FSegmentCount: Integer;
    FNameIndex: DWORD;
    function GetLine(const Idx: Integer): TJclTD32LineInfo;
    function GetLineCount: Integer;
    function GetSegment(const Idx: Integer): TOffsetPair;
  public
    constructor Create(pSrcFile: PSourceFileEntry; Base: NativeInt);
    destructor Destroy; override;
    function FindLine(const AAddr: DWORD; out ALine: TJclTD32LineInfo): Boolean;
    property NameIndex: DWORD read FNameIndex;
    property LineCount: Integer read GetLineCount;
    property Line[const Idx: Integer]: TJclTD32LineInfo read GetLine; default;
    property SegmentCount: Integer read FSegmentCount; //GetSegmentCount;
    property Segment[const Idx: Integer]: TOffsetPair read GetSegment;
  end;

  TJclTD32SymbolInfo = class(TObject)
  private
    FSymbolType: Word;
  public
    constructor Create(pSymInfo: PSymbolInfo); virtual;
    property SymbolType: Word read FSymbolType;
  end;

  TJclTD32ProcSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FNameIndex: DWORD;
    FOffset: DWORD;
    FSize: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property Offset: DWORD read FOffset;
    property Size: DWORD read FSize;
  end;

  TJclTD32LocalProcSymbolInfo = class(TJclTD32ProcSymbolInfo);
  TJclTD32GlobalProcSymbolInfo = class(TJclTD32ProcSymbolInfo);

  { not used by Delphi }
  TJclTD32ObjNameSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FSignature: DWORD;
    FNameIndex: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property Signature: DWORD read FSignature;
  end;

  TJclTD32DataSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FOffset: DWORD;
    FTypeIndex: DWORD;
    FNameIndex: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property TypeIndex: DWORD read FTypeIndex;
    property Offset: DWORD read FOffset;
  end;

  TJclTD32LDataSymbolInfo = class(TJclTD32DataSymbolInfo);
  TJclTD32GDataSymbolInfo = class(TJclTD32DataSymbolInfo);
  TJclTD32PublicSymbolInfo = class(TJclTD32DataSymbolInfo);

  TJclTD32WithSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FOffset: DWORD;
    FSize: DWORD;
    FNameIndex: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property Offset: DWORD read FOffset;
    property Size: DWORD read FSize;
  end;

  { not used by Delphi }
  TJclTD32LabelSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FOffset: DWORD;
    FNameIndex: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property Offset: DWORD read FOffset;
  end;

  { not used by Delphi }
  TJclTD32ConstantSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FValue: DWORD;
    FTypeIndex: DWORD;
    FNameIndex: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property TypeIndex: DWORD read FTypeIndex; // for enums
    property Value: DWORD read FValue;
  end;

  TJclTD32UdtSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FTypeIndex: DWORD;
    FNameIndex: DWORD;
    FProperties: Word;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property NameIndex: DWORD read FNameIndex;
    property TypeIndex: DWORD read FTypeIndex;
    property Properties: Word read FProperties;
  end;

  { not used by Delphi }
  TJclTD32VftPathSymbolInfo = class(TJclTD32SymbolInfo)
  private
    FRootIndex: DWORD;
    FPathIndex: DWORD;
    FOffset: DWORD;
  public
    constructor Create(pSymInfo: PSymbolInfo); override;
    property RootIndex: DWORD read FRootIndex;
    property PathIndex: DWORD read FPathIndex;
    property Offset: DWORD read FOffset;
  end;

  // TD32 parser
  TJclTD32InfoParser = class(TObject)
  private
    FBase: Pointer;
    FData: TCustomMemoryStream;
    FNames: TList;
    FModules: TObjectList;
    FSourceModules: TObjectList;
    FSymbols: TObjectList;
    FProcSymbols: TList;
    FValidData: Boolean;
    function GetName(const Idx: Integer): string;
    function GetNameCount: Integer;
    function GetSymbol(const Idx: Integer): TJclTD32SymbolInfo;
    function GetSymbolCount: Integer;
    function GetProcSymbol(const Idx: Integer): TJclTD32ProcSymbolInfo;
    function GetProcSymbolCount: Integer;
    function GetModule(const Idx: Integer): TJclTD32ModuleInfo;
    function GetModuleCount: Integer;
    function GetSourceModule(const Idx: Integer): TJclTD32SourceModuleInfo;
    function GetSourceModuleCount: Integer;
  protected
    procedure Analyse;
    procedure AnalyseNames(const pSubsection: Pointer; const Size: DWORD); virtual;
    procedure AnalyseGlobalTypes(const pTypes: Pointer; const Size: DWORD); virtual;
    procedure AnalyseAlignSymbols(pSymbols: PSymbolInfos; const Size: DWORD); virtual;
    procedure AnalyseModules(pModInfo: PModuleInfo; const Size: DWORD); virtual;
    procedure AnalyseSourceModules(pSrcModInfo: PSourceModuleInfo; const Size: DWORD); virtual;
    procedure AnalyseUnknownSubSection(const pSubsection: Pointer; const Size: DWORD); virtual;
    function LfaToVa(Lfa: DWORD): Pointer;
  public
    constructor Create(const ATD32Data: TCustomMemoryStream); // Data mustn't be freed before the class is destroyed
    destructor Destroy; override;
    function FindModule(const AAddr: DWORD; out AMod: TJclTD32ModuleInfo): Boolean;
    function FindSourceModule(const AAddr: DWORD; out ASrcMod: TJclTD32SourceModuleInfo): Boolean;
    function FindProc(const AAddr: DWORD; out AProc: TJclTD32ProcSymbolInfo): Boolean;
    class function IsTD32Sign(const Sign: TJclTD32FileSignature): Boolean;
    class function IsTD32DebugInfoValid(const DebugData: Pointer; const DebugDataSize: LongWord): Boolean;
    property Data: TCustomMemoryStream read FData;
    property Names[const Idx: Integer]: string read GetName;
    property NameCount: Integer read GetNameCount;
    property Symbols[const Idx: Integer]: TJclTD32SymbolInfo read GetSymbol;
    property SymbolCount: Integer read GetSymbolCount;
    property ProcSymbols[const Idx: Integer]: TJclTD32ProcSymbolInfo read GetProcSymbol;
    property ProcSymbolCount: Integer read GetProcSymbolCount;
    property Modules[const Idx: Integer]: TJclTD32ModuleInfo read GetModule;
    property ModuleCount: Integer read GetModuleCount;
    property SourceModules[const Idx: Integer]: TJclTD32SourceModuleInfo read GetSourceModule;
    property SourceModuleCount: Integer read GetSourceModuleCount;
    property ValidData: Boolean read FValidData;
  end;

  // TD32 scanner with source location methods
  TJclTD32InfoScanner = class(TJclTD32InfoParser)
  public
    function LineNumberFromAddr(AAddr: DWORD; out Offset: Integer): Integer; overload;
    function LineNumberFromAddr(AAddr: DWORD): Integer; overload;
    function ProcNameFromAddr(AAddr: DWORD): string; overload;
    function ProcNameFromAddr(AAddr: DWORD; out Offset: Integer): string; overload;
    function ModuleNameFromAddr(AAddr: DWORD): string;
    function SourceNameFromAddr(AAddr: DWORD): string;
  end;

type
  TJclReferenceMemoryStream = class(TCustomMemoryStream)
  public
    constructor Create(const Ptr: Pointer; Size: Longint);
    function Write(const Buffer; Count: Longint): Longint; override;
  end;


type
  // Base list
  EJclPeImageError = class(Exception);

  TJclPeImage = class;

  TJclPeImageClass = class of TJclPeImage;

  TJclPeImageBaseList = class(TObjectList)
  private
    FImage: TJclPeImage;
  public
    constructor Create(AImage: TJclPeImage);
    property Image: TJclPeImage read FImage;
  end;

  // Images cache
  TJclPeImagesCache = class(TObject)
  private
    FList: TStringList;
    function GetCount: Integer;
    function GetImages(const FileName: TFileName): TJclPeImage;
  protected
    function GetPeImageClass: TJclPeImageClass; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    property Images[const FileName: TFileName]: TJclPeImage read GetImages; default;
    property Count: Integer read GetCount;
  end;

  // Import section related classes
  TJclPeImportSort = (isName, isOrdinal, isHint, isLibImport);
  TJclPeImportLibSort = (ilName, ilIndex);
  TJclPeImportKind = (ikImport, ikDelayImport, ikBoundImport);
  TJclPeResolveCheck = (icNotChecked, icResolved, icUnresolved);
  TJclPeLinkerProducer = (lrBorland, lrMicrosoft);
  // lrBorland   -> Delphi PE files
  // lrMicrosoft -> MSVC and BCB PE files

  TJclPeImportLibItem = class;

  // Created from a IMAGE_THUNK_DATA64 or IMAGE_THUNK_DATA32 record
  TJclPeImportFuncItem = class(TObject)
  private
    FOrdinal: Word;  // word in 32/64
    FHint: Word;
    FImportLib: TJclPeImportLibItem;
    FIndirectImportName: Boolean;
    FName: string;
    FResolveCheck: TJclPeResolveCheck;
    function GetIsByOrdinal: Boolean;
  protected
    procedure SetName(const Value: string);
    procedure SetIndirectImportName(const Value: string);
    procedure SetResolveCheck(Value: TJclPeResolveCheck);
  public
    constructor Create(AImportLib: TJclPeImportLibItem; AOrdinal: Word;
      AHint: Word; const AName: string);
    property Ordinal: Word read FOrdinal;
    property Hint: Word read FHint;
    property ImportLib: TJclPeImportLibItem read FImportLib;
    property IndirectImportName: Boolean read FIndirectImportName;
    property IsByOrdinal: Boolean read GetIsByOrdinal;
    property Name: string read FName;
    property ResolveCheck: TJclPeResolveCheck read FResolveCheck;
  end;

  // Created from a IMAGE_IMPORT_DESCRIPTOR
  TJclPeImportLibItem = class(TJclPeImageBaseList)
  private
    FImportDescriptor: Pointer;
    FImportDirectoryIndex: Integer;
    FImportKind: TJclPeImportKind;
    FLastSortType: TJclPeImportSort;
    FLastSortDescending: Boolean;
    FName: string;
    FSorted: Boolean;
    FTotalResolveCheck: TJclPeResolveCheck;
    FThunk: Pointer;
    FThunkData: Pointer;
    function GetCount: Integer;
    function GetFileName: TFileName;
    function GetItems(Index: Integer): TJclPeImportFuncItem;
    function GetName: string;
    function GetThunkData32: PImageThunkData32;
    function GetThunkData64: PImageThunkData64;
  protected
    procedure CheckImports(ExportImage: TJclPeImage);
    procedure CreateList;
    procedure SetImportDirectoryIndex(Value: Integer);
    procedure SetImportKind(Value: TJclPeImportKind);
    procedure SetSorted(Value: Boolean);
    procedure SetThunk(Value: Pointer);
  public
    constructor Create(AImage: TJclPeImage; AImportDescriptor: Pointer;
      AImportKind: TJclPeImportKind; const AName: string; AThunk: Pointer);
    procedure SortList(SortType: TJclPeImportSort; Descending: Boolean = False);
    property Count: Integer read GetCount;
    property FileName: TFileName read GetFileName;
    property ImportDescriptor: Pointer read FImportDescriptor;
    property ImportDirectoryIndex: Integer read FImportDirectoryIndex;
    property ImportKind: TJclPeImportKind read FImportKind;
    property Items[Index: Integer]: TJclPeImportFuncItem read GetItems; default;
    property Name: string read GetName;
    property OriginalName: string read FName;
    // use the following properties
    // property ThunkData: PImageThunkData
    property ThunkData32: PImageThunkData32 read GetThunkData32;
    property ThunkData64: PImageThunkData64 read GetThunkData64;
    property TotalResolveCheck: TJclPeResolveCheck read FTotalResolveCheck;
  end;

  TJclPeImportList = class(TJclPeImageBaseList)
  private
    FAllItemsList: TList;
    FFilterModuleName: string;
    FLastAllSortType: TJclPeImportSort;
    FLastAllSortDescending: Boolean;
    FLinkerProducer: TJclPeLinkerProducer;
    FParallelImportTable: array of Pointer;
    FUniqueNamesList: TStringList;
    function GetAllItemCount: Integer;
    function GetAllItems(Index: Integer): TJclPeImportFuncItem;
    function GetItems(Index: Integer): TJclPeImportLibItem;
    function GetUniqueLibItemCount: Integer;
    function GetUniqueLibItems(Index: Integer): TJclPeImportLibItem;
    function GetUniqueLibNames(Index: Integer): string;
    function GetUniqueLibItemFromName(const Name: string): TJclPeImportLibItem;
    procedure SetFilterModuleName(const Value: string);
  protected
    procedure CreateList;
    procedure RefreshAllItems;
  public
    constructor Create(AImage: TJclPeImage);
    destructor Destroy; override;
    procedure CheckImports(PeImageCache: TJclPeImagesCache = nil);
    function MakeBorlandImportTableForMappedImage: Boolean;
    function SmartFindName(const CompareName, LibName: string; Options: TJclSmartCompOptions = []): TJclPeImportFuncItem;
    procedure SortAllItemsList(SortType: TJclPeImportSort; Descending: Boolean = False);
    procedure SortList(SortType: TJclPeImportLibSort);
    procedure TryGetNamesForOrdinalImports;
    property AllItems[Index: Integer]: TJclPeImportFuncItem read GetAllItems;
    property AllItemCount: Integer read GetAllItemCount;
    property FilterModuleName: string read FFilterModuleName write SetFilterModuleName;
    property Items[Index: Integer]: TJclPeImportLibItem read GetItems; default;
    property LinkerProducer: TJclPeLinkerProducer read FLinkerProducer;
    property UniqueLibItemCount: Integer read GetUniqueLibItemCount;
    property UniqueLibItemFromName[const Name: string]: TJclPeImportLibItem read GetUniqueLibItemFromName;
    property UniqueLibItems[Index: Integer]: TJclPeImportLibItem read GetUniqueLibItems;
    property UniqueLibNames[Index: Integer]: string read GetUniqueLibNames;
  end;

  // Export section related classes
  TJclPeExportSort = (esName, esOrdinal, esHint, esAddress, esForwarded,  esAddrOrFwd, esSection);

  TJclPeExportFuncList = class;

  // Created from a IMAGE_EXPORT_DIRECTORY
  TJclPeExportFuncItem = class(TObject)
  private
    FAddress: DWORD;
    FExportList: TJclPeExportFuncList;
    FForwardedName: string;
    FForwardedDotPos: string;
    FHint: Word;
    FName: string;
    FOrdinal: Word;
    FResolveCheck: TJclPeResolveCheck;
    function GetAddressOrForwardStr: string;
    function GetForwardedFuncName: string;
    function GetForwardedLibName: string;
    function GetForwardedFuncOrdinal: DWORD;
    function GetIsExportedVariable: Boolean;
    function GetIsForwarded: Boolean;
    function GetSectionName: string;
    function GetMappedAddress: Pointer;
  protected
    procedure SetResolveCheck(Value: TJclPeResolveCheck);
  public
    constructor Create(AExportList: TJclPeExportFuncList; const AName, AForwardedName: string;
      AAddress: DWORD; AHint: Word; AOrdinal: Word; AResolveCheck: TJclPeResolveCheck);
    property Address: DWORD read FAddress;
    property AddressOrForwardStr: string read GetAddressOrForwardStr;
    property IsExportedVariable: Boolean read GetIsExportedVariable;
    property IsForwarded: Boolean read GetIsForwarded;
    property ForwardedName: string read FForwardedName;
    property ForwardedLibName: string read GetForwardedLibName;
    property ForwardedFuncOrdinal: DWORD read GetForwardedFuncOrdinal;
    property ForwardedFuncName: string read GetForwardedFuncName;
    property Hint: Word read FHint;
    property MappedAddress: Pointer read GetMappedAddress;
    property Name: string read FName;
    property Ordinal: Word read FOrdinal;
    property ResolveCheck: TJclPeResolveCheck read FResolveCheck;
    property SectionName: string read GetSectionName;
  end;

  TJclPeExportFuncList = class(TJclPeImageBaseList)
  private
    FAnyForwards: Boolean;
    FBase: DWORD;
    FExportDir: PImageExportDirectory;
    FForwardedLibsList: TStringList;
    FFunctionCount: DWORD;
    FLastSortType: TJclPeExportSort;
    FLastSortDescending: Boolean;
    FSorted: Boolean;
    FTotalResolveCheck: TJclPeResolveCheck;
    function GetForwardedLibsList: TStrings;
    function GetItems(Index: Integer): TJclPeExportFuncItem;
    function GetItemFromAddress(Address: DWORD): TJclPeExportFuncItem;
    function GetItemFromOrdinal(Ordinal: DWORD): TJclPeExportFuncItem;
    function GetItemFromName(const Name: string): TJclPeExportFuncItem;
    function GetName: string;
  protected
    function CanPerformFastNameSearch: Boolean;
    procedure CreateList;
    property LastSortType: TJclPeExportSort read FLastSortType;
    property LastSortDescending: Boolean read FLastSortDescending;
    property Sorted: Boolean read FSorted;
  public
    constructor Create(AImage: TJclPeImage);
    destructor Destroy; override;
    procedure CheckForwards(PeImageCache: TJclPeImagesCache = nil);
    class function ItemName(Item: TJclPeExportFuncItem): string;
    function OrdinalValid(Ordinal: DWORD): Boolean;
    procedure PrepareForFastNameSearch;
    function SmartFindName(const CompareName: string; Options: TJclSmartCompOptions = []): TJclPeExportFuncItem;
    procedure SortList(SortType: TJclPeExportSort; Descending: Boolean = False);
    property AnyForwards: Boolean read FAnyForwards;
    property Base: DWORD read FBase;
    property ExportDir: PImageExportDirectory read FExportDir;
    property ForwardedLibsList: TStrings read GetForwardedLibsList;
    property FunctionCount: DWORD read FFunctionCount;
    property Items[Index: Integer]: TJclPeExportFuncItem read GetItems; default;
    property ItemFromAddress[Address: DWORD]: TJclPeExportFuncItem read GetItemFromAddress;
    property ItemFromName[const Name: string]: TJclPeExportFuncItem read GetItemFromName;
    property ItemFromOrdinal[Ordinal: DWORD]: TJclPeExportFuncItem read GetItemFromOrdinal;
    property Name: string read GetName;
    property TotalResolveCheck: TJclPeResolveCheck read FTotalResolveCheck;
  end;

  // Resource section related classes
  TJclPeResourceKind = (
    rtUnknown0,
    rtCursorEntry,
    rtBitmap,
    rtIconEntry,
    rtMenu,
    rtDialog,
    rtString,
    rtFontDir,
    rtFont,
    rtAccelerators,
    rtRCData,
    rtMessageTable,
    rtCursor,
    rtUnknown13,
    rtIcon,
    rtUnknown15,
    rtVersion,
    rtDlgInclude,
    rtUnknown18,
    rtPlugPlay,
    rtVxd,
    rtAniCursor,
    rtAniIcon,
    rtHmtl,
    rtManifest,
    rtUserDefined);

  TJclPeResourceList = class;
  TJclPeResourceItem = class;

  TJclPeResourceRawStream = class(TCustomMemoryStream)
  public
    constructor Create(AResourceItem: TJclPeResourceItem);
    function Write(const Buffer; Count: Longint): Longint; override;
  end;

  TJclPeResourceItem = class(TObject)
  private
    FEntry: PImageResourceDirectoryEntry;
    FImage: TJclPeImage;
    FList: TJclPeResourceList;
    FLevel: Byte;
    FParentItem: TJclPeResourceItem;
    FNameCache: string;
    function GetDataEntry: PImageResourceDataEntry;
    function GetIsDirectory: Boolean;
    function GetIsName: Boolean;
    function GetLangID: LANGID;
    function GetList: TJclPeResourceList;
    function GetName: string;
    function GetParameterName: string;
    function GetRawEntryData: Pointer;
    function GetRawEntryDataSize: Integer;
    function GetResourceType: TJclPeResourceKind;
    function GetResourceTypeStr: string;
  protected
    function OffsetToRawData(Ofs: DWORD): NativeInt;
    function Level1Item: TJclPeResourceItem;
    function SubDirData: PImageResourceDirectory;
  public
    constructor Create(AImage: TJclPeImage; AParentItem: TJclPeResourceItem;
      AEntry: PImageResourceDirectoryEntry);
    destructor Destroy; override;
    function CompareName(AName: PChar): Boolean;
    property DataEntry: PImageResourceDataEntry read GetDataEntry;
    property Entry: PImageResourceDirectoryEntry read FEntry;
    property Image: TJclPeImage read FImage;
    property IsDirectory: Boolean read GetIsDirectory;
    property IsName: Boolean read GetIsName;
    property LangID: LANGID read GetLangID;
    property List: TJclPeResourceList read GetList;
    property Level: Byte read FLevel;
    property Name: string read GetName;
    property ParameterName: string read GetParameterName;
    property ParentItem: TJclPeResourceItem read FParentItem;
    property RawEntryData: Pointer read GetRawEntryData;
    property RawEntryDataSize: Integer read GetRawEntryDataSize;
    property ResourceType: TJclPeResourceKind read GetResourceType;
    property ResourceTypeStr: string read GetResourceTypeStr;
  end;

  TJclPeResourceList = class(TJclPeImageBaseList)
  private
    FDirectory: PImageResourceDirectory;
    FParentItem: TJclPeResourceItem;
    function GetItems(Index: Integer): TJclPeResourceItem;
  protected
    procedure CreateList(AParentItem: TJclPeResourceItem);
  public
    constructor Create(AImage: TJclPeImage; AParentItem: TJclPeResourceItem;
      ADirectory: PImageResourceDirectory);
    function FindName(const Name: string): TJclPeResourceItem;
    property Directory: PImageResourceDirectory read FDirectory;
    property Items[Index: Integer]: TJclPeResourceItem read GetItems; default;
    property ParentItem: TJclPeResourceItem read FParentItem;
  end;

  TJclPeRootResourceList = class(TJclPeResourceList)
  private
    FManifestContent: TStringList;
    function GetManifestContent: TStrings;
  public
    destructor Destroy; override;
    function FindResource(ResourceType: TJclPeResourceKind;
      const ResourceName: string = ''): TJclPeResourceItem; overload;
    function FindResource(const ResourceType: PChar;
      const ResourceName: PChar = nil): TJclPeResourceItem; overload;
    function ListResourceNames(ResourceType: TJclPeResourceKind; const Strings: TStrings): Boolean;
    property ManifestContent: TStrings read GetManifestContent;
  end;

  // Relocation section related classes
  TJclPeRelocation = record
    Address: Word;
    RelocType: Byte;
    VirtualAddress: DWORD;
  end;

  TJclPeRelocEntry = class(TObject)
  private
    FChunk: PImageBaseRelocation;
    FCount: Integer;
    function GetRelocations(Index: Integer): TJclPeRelocation;
    function GetSize: DWORD;
    function GetVirtualAddress: DWORD;
  public
    constructor Create(AChunk: PImageBaseRelocation; ACount: Integer);
    property Count: Integer read FCount;
    property Relocations[Index: Integer]: TJclPeRelocation read GetRelocations; default;
    property Size: DWORD read GetSize;
    property VirtualAddress: DWORD read GetVirtualAddress;
  end;

  TJclPeRelocList = class(TJclPeImageBaseList)
  private
    FAllItemCount: Integer;
    function GetItems(Index: Integer): TJclPeRelocEntry;
    function GetAllItems(Index: Integer): TJclPeRelocation;
  protected
    procedure CreateList;
  public
    constructor Create(AImage: TJclPeImage);
    property AllItems[Index: Integer]: TJclPeRelocation read GetAllItems;
    property AllItemCount: Integer read FAllItemCount;
    property Items[Index: Integer]: TJclPeRelocEntry read GetItems; default;
  end;

  // Debug section related classes
  TJclPeDebugList = class(TJclPeImageBaseList)
  private
    function GetItems(Index: Integer): TImageDebugDirectory;
  protected
    procedure CreateList;
  public
    constructor Create(AImage: TJclPeImage);
    property Items[Index: Integer]: TImageDebugDirectory read GetItems; default;
  end;

  // Certificates section related classes
  TJclPeCertificate = class(TObject)
  private
    FData: Pointer;
    FHeader: TWinCertificate;
  public
    constructor Create(AHeader: TWinCertificate; AData: Pointer);
    property Data: Pointer read FData;
    property Header: TWinCertificate read FHeader;
  end;

  TJclPeCertificateList = class(TJclPeImageBaseList)
  private
    function GetItems(Index: Integer): TJclPeCertificate;
  protected
    procedure CreateList;
  public
    constructor Create(AImage: TJclPeImage);
    property Items[Index: Integer]: TJclPeCertificate read GetItems; default;
  end;

  // Common Language Runtime section related classes
  TJclPeCLRHeader = class(TObject)
  private
    FHeader: TImageCor20Header;
    FImage: TJclPeImage;
    function GetVersionString: string;
    function GetHasMetadata: Boolean;
  protected
    procedure ReadHeader;
  public
    constructor Create(AImage: TJclPeImage);
    property HasMetadata: Boolean read GetHasMetadata;
    property Header: TImageCor20Header read FHeader;
    property VersionString: string read GetVersionString;
    property Image: TJclPeImage read FImage;
  end;

  // PE Image
  TJclPeHeader = (
    JclPeHeader_Signature,
    JclPeHeader_Machine,
    JclPeHeader_NumberOfSections,
    JclPeHeader_TimeDateStamp,
    JclPeHeader_PointerToSymbolTable,
    JclPeHeader_NumberOfSymbols,
    JclPeHeader_SizeOfOptionalHeader,
    JclPeHeader_Characteristics,
    JclPeHeader_Magic,
    JclPeHeader_LinkerVersion,
    JclPeHeader_SizeOfCode,
    JclPeHeader_SizeOfInitializedData,
    JclPeHeader_SizeOfUninitializedData,
    JclPeHeader_AddressOfEntryPoint,
    JclPeHeader_BaseOfCode,
    JclPeHeader_BaseOfData,
    JclPeHeader_ImageBase,
    JclPeHeader_SectionAlignment,
    JclPeHeader_FileAlignment,
    JclPeHeader_OperatingSystemVersion,
    JclPeHeader_ImageVersion,
    JclPeHeader_SubsystemVersion,
    JclPeHeader_Win32VersionValue,
    JclPeHeader_SizeOfImage,
    JclPeHeader_SizeOfHeaders,
    JclPeHeader_CheckSum,
    JclPeHeader_Subsystem,
    JclPeHeader_DllCharacteristics,
    JclPeHeader_SizeOfStackReserve,
    JclPeHeader_SizeOfStackCommit,
    JclPeHeader_SizeOfHeapReserve,
    JclPeHeader_SizeOfHeapCommit,
    JclPeHeader_LoaderFlags,
    JclPeHeader_NumberOfRvaAndSizes);

  TJclLoadConfig = (
    JclLoadConfig_Characteristics,   { TODO : rename to Size? }
    JclLoadConfig_TimeDateStamp,
    JclLoadConfig_Version,
    JclLoadConfig_GlobalFlagsClear,
    JclLoadConfig_GlobalFlagsSet,
    JclLoadConfig_CriticalSectionDefaultTimeout,
    JclLoadConfig_DeCommitFreeBlockThreshold,
    JclLoadConfig_DeCommitTotalFreeThreshold,
    JclLoadConfig_LockPrefixTable,
    JclLoadConfig_MaximumAllocationSize,
    JclLoadConfig_VirtualMemoryThreshold,
    JclLoadConfig_ProcessHeapFlags,
    JclLoadConfig_ProcessAffinityMask,
    JclLoadConfig_CSDVersion,
    JclLoadConfig_Reserved1,
    JclLoadConfig_EditList,
    JclLoadConfig_Reserved           { TODO : extend to the new fields? }
  );

  TJclPeImageStatus = (stNotLoaded, stOk, stNotPE, stNotSupported, stNotFound, stError);
  TJclPeTarget = (taUnknown, taWin32, taWin64);

  TJclPeImage = class(TObject)
  private
    FAttachedImage: Boolean;
    FCertificateList: TJclPeCertificateList;
    FCLRHeader: TJclPeCLRHeader;
    FDebugList: TJclPeDebugList;
    FFileName: TFileName;
    FImageSections: TStringList;
    FLoadedImage: TLoadedImage;
    FExportList: TJclPeExportFuncList;
    FImportList: TJclPeImportList;
    FNoExceptions: Boolean;
    FReadOnlyAccess: Boolean;
    FRelocationList: TJclPeRelocList;
    FResourceList: TJclPeRootResourceList;
    FResourceVA: NativeInt;
    FStatus: TJclPeImageStatus;
    FTarget: TJclPeTarget;
    FStringTable: TStringList;
    function GetCertificateList: TJclPeCertificateList;
    function GetCLRHeader: TJclPeCLRHeader;
    function GetDebugList: TJclPeDebugList;
    function GetDescription: string;
    function GetDirectories(Directory: Word): TImageDataDirectory;
    function GetDirectoryExists(Directory: Word): Boolean;
    function GetExportList: TJclPeExportFuncList;
    function GetImageSectionCount: Integer;
    function GetImageSectionHeaders(Index: Integer): TImageSectionHeader;
    function GetImageSectionNames(Index: Integer): string;
    function GetImageSectionNameFromRva(const Rva: DWORD): string;
    function GetImportList: TJclPeImportList;
    function GetLoadConfigValues(Index: TJclLoadConfig): string;
    function GetMappedAddress: NativeInt;
    function GetOptionalHeader32: TImageOptionalHeader32;
    function GetOptionalHeader64: TImageOptionalHeader64;
    function GetRelocationList: TJclPeRelocList;
    function GetResourceList: TJclPeRootResourceList;
    function GetUnusedHeaderBytes: TImageDataDirectory;
    function GetVersionInfoAvailable: Boolean;
    procedure ReadImageSections;
    procedure ReadStringTable;
    procedure SetFileName(const Value: TFileName);
    function GetStringTableCount: Integer;
    function GetStringTableItem(Index: Integer): string;
    function GetImageSectionFullNames(Index: Integer): string;
  protected
    procedure AfterOpen; dynamic;
    procedure CheckNotAttached;
    procedure Clear; dynamic;
    function ExpandModuleName(const ModuleName: string): TFileName;
    procedure RaiseStatusException;
    function ResourceItemCreate(AEntry: PImageResourceDirectoryEntry;
      AParentItem: TJclPeResourceItem): TJclPeResourceItem; virtual;
    function ResourceListCreate(ADirectory: PImageResourceDirectory;
      AParentItem: TJclPeResourceItem): TJclPeResourceList; virtual;
    property NoExceptions: Boolean read FNoExceptions;
  public
    constructor Create(ANoExceptions: Boolean = False); virtual;
    destructor Destroy; override;
    procedure AttachLoadedModule(const Handle: HMODULE);
    function CalculateCheckSum: DWORD;
    function DirectoryEntryToData(Directory: Word): Pointer;
    function GetSectionHeader(const SectionName: string; out Header: PImageSectionHeader): Boolean;
    function GetSectionName(Header: PImageSectionHeader): string;
    function GetNameInStringTable(Offset: ULONG): string;
    function IsBrokenFormat: Boolean;
    function IsCLR: Boolean;
    function IsSystemImage: Boolean;
    // RVA are always DWORD
    function RawToVa(Raw: DWORD): Pointer; overload;
    function RvaToSection(Rva: DWORD): PImageSectionHeader; overload;
    function RvaToVa(Rva: DWORD): Pointer; overload;
    function RvaToVaEx(Rva: DWORD): Pointer; overload;
    function StatusOK: Boolean;
    procedure TryGetNamesForOrdinalImports;
    function VerifyCheckSum: Boolean;
    class function DebugTypeNames(DebugType: DWORD): string;
    //class function DirectoryNames(Directory: Word): string;
    class function ExpandBySearchPath(const ModuleName, BasePath: string): TFileName;
    class function HeaderNames(Index: TJclPeHeader): string;
    class function LoadConfigNames(Index: TJclLoadConfig): string;
    class function ShortSectionInfo(Characteristics: DWORD): string;
    class function DateTimeToStamp(const DateTime: TDateTime): DWORD;
    class function StampToDateTime(TimeDateStamp: DWORD): TDateTime;
    property AttachedImage: Boolean read FAttachedImage;
    property CertificateList: TJclPeCertificateList read GetCertificateList;
    property CLRHeader: TJclPeCLRHeader read GetCLRHeader;
    property DebugList: TJclPeDebugList read GetDebugList;
    property Description: string read GetDescription;
    property Directories[Directory: Word]: TImageDataDirectory read GetDirectories;
    property DirectoryExists[Directory: Word]: Boolean read GetDirectoryExists;
    property ExportList: TJclPeExportFuncList read GetExportList;
    property FileName: TFileName read FFileName write SetFileName;
    property ImageSectionCount: Integer read GetImageSectionCount;
    property ImageSectionHeaders[Index: Integer]: TImageSectionHeader read GetImageSectionHeaders;
    property ImageSectionNames[Index: Integer]: string read GetImageSectionNames;
    property ImageSectionFullNames[Index: Integer]: string read GetImageSectionFullNames;
    property ImageSectionNameFromRva[const Rva: DWORD]: string read GetImageSectionNameFromRva;
    property ImportList: TJclPeImportList read GetImportList;
    property LoadConfigValues[Index: TJclLoadConfig]: string read GetLoadConfigValues;
    property LoadedImage: TLoadedImage read FLoadedImage;
    property MappedAddress: NativeInt read GetMappedAddress;
    property StringTableCount: Integer read GetStringTableCount;
    property StringTable[Index: Integer]: string read GetStringTableItem;
    // use the following properties
    property OptionalHeader32: TImageOptionalHeader32 read GetOptionalHeader32;
    property OptionalHeader64: TImageOptionalHeader64 read GetOptionalHeader64;
    property ReadOnlyAccess: Boolean read FReadOnlyAccess write FReadOnlyAccess;
    property RelocationList: TJclPeRelocList read GetRelocationList;
    property ResourceVA: NativeInt read FResourceVA;
    property ResourceList: TJclPeRootResourceList read GetResourceList;
    property Status: TJclPeImageStatus read FStatus;
    property Target: TJclPeTarget read FTarget;
    property UnusedHeaderBytes: TImageDataDirectory read GetUnusedHeaderBytes;
    property VersionInfoAvailable: Boolean read GetVersionInfoAvailable;
  end;


  // Borland Delphi PE Image specific information
  TJclPePackageInfo = class(TObject)
  private
    FAvailable: Boolean;
    FContains: TStringList;
    FDcpName: string;
    FRequires: TStringList;
    FFlags: Integer;
    FDescription: string;
    FEnsureExtension: Boolean;
    FSorted: Boolean;
    function GetContains: TStrings;
    function GetContainsCount: Integer;
    function GetContainsFlags(Index: Integer): Byte;
    function GetContainsNames(Index: Integer): string;
    function GetRequires: TStrings;
    function GetRequiresCount: Integer;
    function GetRequiresNames(Index: Integer): string;
  protected
    procedure ReadPackageInfo(ALibHandle: THandle);
    procedure SetDcpName(const Value: string);
  public
    constructor Create(ALibHandle: THandle);
    destructor Destroy; override;
    property Available: Boolean read FAvailable;
    property Contains: TStrings read GetContains;
    property ContainsCount: Integer read GetContainsCount;
    property ContainsNames[Index: Integer]: string read GetContainsNames;
    property ContainsFlags[Index: Integer]: Byte read GetContainsFlags;
    property Description: string read FDescription;
    property DcpName: string read FDcpName;
    property EnsureExtension: Boolean read FEnsureExtension write FEnsureExtension;
    property Flags: Integer read FFlags;
    property Requires: TStrings read GetRequires;
    property RequiresCount: Integer read GetRequiresCount;
    property RequiresNames[Index: Integer]: string read GetRequiresNames;
    property Sorted: Boolean read FSorted write FSorted;
  end;

  TJclPeBorForm = class(TObject)
  private
    FFormFlags: TFilerFlags;
    FFormClassName: string;
    FFormObjectName: string;
    FFormPosition: Integer;
    FResItem: TJclPeResourceItem;
    function GetDisplayName: string;
  public
    constructor Create(AResItem: TJclPeResourceItem; AFormFlags: TFilerFlags;
      AFormPosition: Integer; const AFormClassName, AFormObjectName: string);
    procedure ConvertFormToText(const Stream: TStream); overload;
    procedure ConvertFormToText(const Strings: TStrings); overload;
    property FormClassName: string read FFormClassName;
    property FormFlags: TFilerFlags read FFormFlags;
    property FormObjectName: string read FFormObjectName;
    property FormPosition: Integer read FFormPosition;
    property DisplayName: string read GetDisplayName;
    property ResItem: TJclPeResourceItem read FResItem;
  end;

  TJclPeBorImage = class;



  TJclPeBorImage = class(TJclPeImage)
  private
    FForms: TObjectList;
    FIsPackage: Boolean;
    FIsBorlandImage: Boolean;
    FLibHandle: THandle;
    FPackageInfo: TJclPePackageInfo;
    FPackageInfoSorted: Boolean;
    FPackageCompilerVersion: Integer;
    function GetFormCount: Integer;
    function GetForms(Index: Integer): TJclPeBorForm;
    function GetFormFromName(const FormClassName: string): TJclPeBorForm;
    function GetLibHandle: THandle;
    function GetPackageCompilerVersion: Integer;
    function GetPackageInfo: TJclPePackageInfo;
  protected
    procedure AfterOpen; override;
    procedure Clear; override;
    procedure CreateFormsList;
  public
    constructor Create(ANoExceptions: Boolean = False); override;
    destructor Destroy; override;
    function DependedPackages(List: TStrings; FullPathName, Descriptions: Boolean): Boolean;
    function FreeLibHandle: Boolean;
    property Forms[Index: Integer]: TJclPeBorForm read GetForms;
    property FormCount: Integer read GetFormCount;
    property FormFromName[const FormClassName: string]: TJclPeBorForm read GetFormFromName;
    property IsBorlandImage: Boolean read FIsBorlandImage;
    property IsPackage: Boolean read FIsPackage;
    property LibHandle: THandle read GetLibHandle;
    property PackageCompilerVersion: Integer read GetPackageCompilerVersion;
    property PackageInfo: TJclPePackageInfo read GetPackageInfo;
    property PackageInfoSorted: Boolean read FPackageInfoSorted write FPackageInfoSorted;
  end;

  // Threaded function search
  TJclPeNameSearchOption = (seImports, seDelayImports, seBoundImports, seExports);
  TJclPeNameSearchOptions = set of TJclPeNameSearchOption;

  TJclPeNameSearchNotifyEvent = procedure (Sender: TObject; PeImage: TJclPeImage;
    var Process: Boolean) of object;
  TJclPeNameSearchFoundEvent = procedure (Sender: TObject; const FileName: TFileName;
    const FunctionName: string; Option: TJclPeNameSearchOption) of object;


// PE Image with TD32 information and source location support
  TJclPeBorTD32Image = class(TJclPeBorImage)
  private
    FIsTD32DebugPresent: Boolean;
    FTD32DebugData: TCustomMemoryStream;
    FTD32Scanner: TJclTD32InfoScanner;
  protected
    procedure AfterOpen; override;
    procedure Clear; override;
    procedure ClearDebugData;
    procedure CheckDebugData;
    function IsDebugInfoInImage(var DataStream: TCustomMemoryStream): Boolean;
    function IsDebugInfoInTds(var DataStream: TCustomMemoryStream): Boolean;
  public
    property IsTD32DebugPresent: Boolean read FIsTD32DebugPresent;
    property TD32DebugData: TCustomMemoryStream read FTD32DebugData;
    property TD32Scanner: TJclTD32InfoScanner read FTD32Scanner;
  end;


// PE Image miscellaneous functions
type
  TJclRebaseImageInfo32 = record
    OldImageSize: DWORD;
    OldImageBase: LongWord;
    NewImageSize: DWORD;
    NewImageBase: LongWord;
  end;
  TJclRebaseImageInfo64 = record
    OldImageSize: DWORD;
    OldImageBase: Int64;
    NewImageSize: DWORD;
    NewImageBase: Int64;
  end;


// Various simple PE Image searching and listing routines
{ Exports searching }

// Mapped or loaded image related routines
// use PeMapImgNtHeaders32
// function PeMapImgNtHeaders(const BaseAddress: Pointer): PImageNtHeaders;
function PeMapImgNtHeaders32(const BaseAddress: Pointer): PImageNtHeaders32; overload;
function PeMapImgNtHeaders64(const BaseAddress: Pointer): PImageNtHeaders64; overload;

function PeMapImgLibraryName(const BaseAddress: Pointer): string;
function PeMapImgLibraryName32(const BaseAddress: Pointer): string;
function PeMapImgLibraryName64(const BaseAddress: Pointer): string;

function PeMapImgSize(const BaseAddress: Pointer): DWORD; overload;
function PeMapImgSize32(const BaseAddress: Pointer): DWORD; overload;
function PeMapImgSize64(const BaseAddress: Pointer): DWORD; overload;

function PeMapImgTarget(const BaseAddress: Pointer): TJclPeTarget; overload;

type
  TImageSectionHeaderArray = array of TImageSectionHeader;

// use PeMapImgSections32
// function PeMapImgSections(NtHeaders: PImageNtHeaders): PImageSectionHeader;
function PeMapImgSections32(NtHeaders: PImageNtHeaders32): PImageSectionHeader; overload;
function PeMapImgSections32(Stream: TStream; const NtHeaders32Position: Int64; const NtHeaders32: TImageNtHeaders32;
  out ImageSectionHeaders: TImageSectionHeaderArray): Int64; overload;
function PeMapImgSections64(NtHeaders: PImageNtHeaders64): PImageSectionHeader; overload;
function PeMapImgSections64(Stream: TStream; const NtHeaders64Position: Int64; const NtHeaders64: TImageNtHeaders64;
  out ImageSectionHeaders: TImageSectionHeaderArray): Int64; overload;

// use PeMapImgFindSection32
// function PeMapImgFindSection(NtHeaders: PImageNtHeaders;
//   const SectionName: string): PImageSectionHeader;
function PeMapImgFindSection32(NtHeaders: PImageNtHeaders32;
  const SectionName: string): PImageSectionHeader;
function PeMapImgFindSection64(NtHeaders: PImageNtHeaders64;
  const SectionName: string): PImageSectionHeader;

function PeMapImgFindSectionFromModule(const BaseAddress: Pointer;
  const SectionName: string): PImageSectionHeader;

function PeMapImgExportedVariables(const Module: HMODULE; const VariablesList: TStrings): Boolean;

function PeMapImgResolvePackageThunk(Address: Pointer): Pointer;

function PeMapFindResource(const Module: HMODULE; const ResourceType: PChar;
  const ResourceName: string): Pointer;

type
  TJclPeSectionStream = class(TCustomMemoryStream)
  private
    FInstance: HMODULE;
    FSectionHeader: TImageSectionHeader;
    procedure Initialize(Instance: HMODULE; const ASectionName: string);
  public
    constructor Create(Instance: HMODULE; const ASectionName: string);
    function Write(const Buffer; Count: Longint): Longint; override;
    property Instance: HMODULE read FInstance;
    property SectionHeader: TImageSectionHeader read FSectionHeader;
  end;

// API hooking classes


// Borland BPL packages name unmangling
type
  TJclBorUmSymbolKind = (skData, skFunction, skConstructor, skDestructor, skRTTI, skVTable);
  TJclBorUmSymbolModifier = (smQualified, smLinkProc);
  TJclBorUmSymbolModifiers = set of TJclBorUmSymbolModifier;
  TJclBorUmDescription = record
    Kind: TJclBorUmSymbolKind;
    Modifiers: TJclBorUmSymbolModifiers;
  end;
  TJclBorUmResult = (urOk, urNotMangled, urMicrosoft, urError);
  TJclPeUmResult = (umNotMangled, umBorland, umMicrosoft);

function UndecorateSymbolName(const DecoratedName: string; out UnMangled: string; Flags: DWORD): Boolean;

type
  NT_TIB32 = packed record
    ExceptionList: DWORD;
    StackBase: DWORD;
    StackLimit: DWORD;
    SubSystemTib: DWORD;
    case Integer of
      0 : (
        FiberData: DWORD;
        ArbitraryUserPointer: DWORD;
        Self: DWORD;
      );
      1 : (
        Version: DWORD;
      );
  end;

// Optimized functionality of JclSysInfo functions ModuleFromAddr and IsSystemModule
type
  TJclModuleInfo = class(TObject)
  private
    FSize: Cardinal;
    FEndAddr: Pointer;
    FStartAddr: Pointer;
    FSystemModule: Boolean;
  public
    property EndAddr: Pointer read FEndAddr;
    property Size: Cardinal read FSize;
    property StartAddr: Pointer read FStartAddr;
    property SystemModule: Boolean read FSystemModule;
  end;

  TJclModuleInfoList = class(TObjectList)
  private
    FDynamicBuild: Boolean;
    FSystemModulesOnly: Boolean;
    function GetItems(Index: Integer): TJclModuleInfo;
    function GetModuleFromAddress(Addr: Pointer): TJclModuleInfo;
  protected
    procedure BuildModulesList;
    function CreateItemForAddress(Addr: Pointer; SystemModule: Boolean): TJclModuleInfo;
  public
    constructor Create(ADynamicBuild, ASystemModulesOnly: Boolean);
    function AddModule(Module: HMODULE; SystemModule: Boolean): Boolean;
    function IsSystemModuleAddress(Addr: Pointer): Boolean;
    function IsValidModuleAddress(Addr: Pointer): Boolean;
    property DynamicBuild: Boolean read FDynamicBuild;
    property Items[Index: Integer]: TJclModuleInfo read GetItems;
    property ModuleFromAddress[Addr: Pointer]: TJclModuleInfo read GetModuleFromAddress;
  end;

function JclValidateModuleAddress(Addr: Pointer): Boolean;

type
  PJclMapLineNumber = ^TJclMapLineNumber;
  TJclMapLineNumber = record
    Segment: Word;
    VA: DWORD; // VA relative to (module base address + $10000)
    LineNumber: Integer;
  end;


type
  PJclDbgHeader = ^TJclDbgHeader;
  TJclDbgHeader = packed record
    Signature: DWORD;
    Version: Byte;
    Units: Integer;
    SourceNames: Integer;
    Symbols: Integer;
    LineNumbers: Integer;
    Words: Integer;
    ModuleName: Integer;
    CheckSum: Integer;
    CheckSumValid: Boolean;
  end;

  TJclBinDbgNameCache = record
    Addr: DWORD;
    FirstWord: Integer;
    SecondWord: Integer;
  end;

  TJclBinDebugScanner = class(TObject)
  private
    FCacheData: Boolean;
    FStream: TCustomMemoryStream;
    FValidFormat: Boolean;
    FLineNumbers: array of TJclMapLineNumber;
    FProcNames: array of TJclBinDbgNameCache;
    function GetModuleName: string;
  protected
    procedure CacheLineNumbers;
    procedure CacheProcNames;
    procedure CheckFormat;
    function DataToStr(A: Integer): string;
    function MakePtr(A: Integer): Pointer;
    function ReadValue(var P: Pointer; var Value: Integer): Boolean;
  public
    constructor Create(AStream: TCustomMemoryStream; CacheData: Boolean);
    function IsModuleNameValid(const Name: TFileName): Boolean;
    function LineNumberFromAddr(Addr: DWORD): Integer; overload;
    function LineNumberFromAddr(Addr: DWORD; out Offset: Integer): Integer; overload;
    function ProcNameFromAddr(Addr: DWORD): string; overload;
    function ProcNameFromAddr(Addr: DWORD; out Offset: Integer): string; overload;
    function ModuleNameFromAddr(Addr: DWORD): string;
    function ModuleStartFromAddr(Addr: DWORD): DWORD;
    function SourceNameFromAddr(Addr: DWORD): string;
    property ModuleName: string read GetModuleName;
    property ValidFormat: Boolean read FValidFormat;
  end;

// Source Locations
type
  TJclDebugInfoSource = class;

  PJclLocationInfo = ^TJclLocationInfo;
  TJclLocationInfo = record
    Address: Pointer;               // Error address
    UnitName: string;               // Name of Delphi unit
    ProcedureName: string;          // Procedure name
    OffsetFromProcName: Integer;    // Offset from Address to ProcedureName symbol location
    LineNumber: Integer;            // Line number
    OffsetFromLineNumber: Integer;  // Offset from Address to LineNumber symbol location
    SourceName: string;             // Module file name
    DebugInfo: TJclDebugInfoSource; // Location object
    BinaryFileName: string;         // Name of the binary file containing the symbol
  end;

  TJclLocationInfoExValues = set of (lievLocationInfo, lievProcedureStartLocationInfo, lievUnitVersionInfo);

  TJclLocationInfoListOptions = set of (liloAutoGetAddressInfo, liloAutoGetLocationInfo, liloAutoGetUnitVersionInfo);


  TJclDebugInfoSource = class(TObject)
  private
    FModule: HMODULE;
    function GetFileName: TFileName;
  protected
    function VAFromAddr(const Addr: Pointer): DWORD; virtual;
  public
    constructor Create(AModule: HMODULE); virtual;
    function InitializeSource: Boolean; virtual; abstract;
    function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean; virtual; abstract;
    property Module: HMODULE read FModule;
    property FileName: TFileName read GetFileName;
  end;

  TJclDebugInfoSourceClass = class of TJclDebugInfoSource;

  TJclDebugInfoList = class(TObjectList)
  private
    function GetItemFromModule(const Module: HMODULE): TJclDebugInfoSource;
    function GetItems(Index: Integer): TJclDebugInfoSource;
  protected
    function CreateDebugInfo(const Module: HMODULE): TJclDebugInfoSource;
  public
    class procedure RegisterDebugInfoSource(
      const InfoSourceClass: TJclDebugInfoSourceClass);
    class procedure UnRegisterDebugInfoSource(
      const InfoSourceClass: TJclDebugInfoSourceClass);
    class procedure RegisterDebugInfoSourceFirst(
      const InfoSourceClass: TJclDebugInfoSourceClass);
    class procedure NeedInfoSourceClassList;
    function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean;
    property ItemFromModule[const Module: HMODULE]: TJclDebugInfoSource read GetItemFromModule;
    property Items[Index: Integer]: TJclDebugInfoSource read GetItems;
  end;


  TJclDebugInfoBinary = class(TJclDebugInfoSource)
  private
    FScanner: TJclBinDebugScanner;
    FStream: TCustomMemoryStream;
  public
    destructor Destroy; override;
    function InitializeSource: Boolean; override;
    function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean; override;
  end;

  TJclDebugInfoTD32 = class(TJclDebugInfoSource)
  private
    FImage: TJclPeBorTD32Image;
  public
    destructor Destroy; override;
    function InitializeSource: Boolean; override;
    function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean; override;
  end;

  TJclDebugInfoSymbols = class(TJclDebugInfoSource)
  public
    class function LoadDebugFunctions: Boolean;
    class function UnloadDebugFunctions: Boolean;
    class function InitializeDebugSymbols: Boolean;
    class function CleanupDebugSymbols: Boolean;
    function InitializeSource: Boolean; override;
    function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean; override;
  end;

// Source location functions
function Caller(Level: Integer = 0; FastStackWalk: Boolean = False): Pointer;

function GetLocationInfo(const Addr: Pointer): TJclLocationInfo; overload;
function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean; overload;
function GetLocationInfoStr(const Addr: Pointer; IncludeModuleName: Boolean = False;
  IncludeAddressOffset: Boolean = False; IncludeStartProcLineOffset: Boolean = False;
  IncludeVAddress: Boolean = False): string;

// Stack info routines base list
type
  TJclStackBaseList = class(TObjectList)
  private
    FThreadID: DWORD;
    FTimeStamp: TDateTime;
  protected
    FOnDestroy: TNotifyEvent;
  public
    constructor Create;
    destructor Destroy; override;
    property ThreadID: DWORD read FThreadID;
    property TimeStamp: TDateTime read FTimeStamp;
  end;


type
  TJclByteArray = array [0..MaxInt div SizeOf(Byte) - 1] of Byte;
  PJclByteArray = ^TJclByteArray;
// Stack info routines
type
  PDWORD_PTRArray = ^TDWORD_PTRArray;
  TDWORD_PTRArray = array [0..(MaxInt - $F) div SizeOf(DWORD_PTR)] of DWORD_PTR;
  PDWORD_PTR = ^DWORD_PTR;

  PStackFrame = ^TStackFrame;
  TStackFrame = record
    CallerFrame: NativeInt;
    CallerAddr: NativeInt;
  end;

  PStackInfo = ^TStackInfo;
  TStackInfo = record
    CallerAddr: NativeInt;
    Level: Integer;
    CallerFrame: NativeInt;
    DumpSize: DWORD;
    ParamSize: DWORD;
    ParamPtr: PDWORD_PTRArray;
    case Integer of
      0:
        (StackFrame: PStackFrame);
      1:
        (DumpPtr: PJclByteArray);
  end;

  TJclStackInfoItem = class(TObject)
  private
    FStackInfo: TStackInfo;
    function GetCallerAddr: Pointer;
    function GetLogicalAddress: NativeInt;
  public
    property CallerAddr: Pointer read GetCallerAddr;
    property LogicalAddress: NativeInt read GetLogicalAddress;
    property StackInfo: TStackInfo read FStackInfo;
  end;

  TJclStackInfoList = class(TJclStackBaseList)
  private
    FIgnoreLevels: Integer;
    TopOfStack: NativeInt;
    BaseOfStack: NativeInt;
    FStackData: PPointer;
    FFramePointer: Pointer;
    FModuleInfoList: TJclModuleInfoList;
    FCorrectOnAccess: Boolean;
    FSkipFirstItem: Boolean;
    FDelayedTrace: Boolean;
    FInStackTracing: Boolean;
    FRaw: Boolean;
    FStackOffset: Int64;
    {$IFDEF WIN64}
    procedure CaptureBackTrace;
    {$ENDIF WIN64}
    function GetItems(Index: Integer): TJclStackInfoItem;
    function NextStackFrame(var StackFrame: PStackFrame; var StackInfo: TStackInfo): Boolean;
    procedure StoreToList(const StackInfo: TStackInfo);
    procedure TraceStackFrames;
    procedure TraceStackRaw;
    {$IFDEF WIN32}
    procedure DelayStoreStack;
    {$ENDIF WIN32}
    function ValidCallSite(CodeAddr: NativeInt; out CallInstructionSize: Cardinal): Boolean;
    function ValidStackAddr(StackAddr: NativeInt): Boolean;
    function GetCount: Integer;
  public
    constructor Create(ARaw: Boolean; AIgnoreLevels: Integer;
      AFirstCaller: Pointer); overload;
    constructor Create(ARaw: Boolean; AIgnoreLevels: Integer;
      AFirstCaller: Pointer; ADelayedTrace: Boolean); overload;
    constructor Create(ARaw: Boolean; AIgnoreLevels: Integer;
      AFirstCaller: Pointer; ADelayedTrace: Boolean; ABaseOfStack: Pointer); overload;
    constructor Create(ARaw: Boolean; AIgnoreLevels: Integer;
      AFirstCaller: Pointer; ADelayedTrace: Boolean; ABaseOfStack, ATopOfStack: Pointer); overload;
    destructor Destroy; override;
    procedure ForceStackTracing;
    procedure AddToStrings(Strings: TStrings; IncludeModuleName: Boolean = False;
      IncludeAddressOffset: Boolean = False; IncludeStartProcLineOffset: Boolean = False;
      IncludeVAddress: Boolean = False);
    property DelayedTrace: Boolean read FDelayedTrace;
    property Items[Index: Integer]: TJclStackInfoItem read GetItems; default;
    property IgnoreLevels: Integer read FIgnoreLevels;
    property Count: Integer read GetCount;
    property Raw: Boolean read FRaw;
  end;

function JclCreateStackList(Raw: Boolean; AIgnoreLevels: Integer; FirstCaller: Pointer): TJclStackInfoList; overload;


function JclLastExceptStackList: TJclStackInfoList;
function JclLastExceptStackListToStrings(Strings: TStrings; IncludeModuleName: Boolean = False;
  IncludeAddressOffset: Boolean = False; IncludeStartProcLineOffset: Boolean = False;
  IncludeVAddress: Boolean = False): Boolean;

function JclGetExceptStackList(ThreadID: DWORD): TJclStackInfoList;
function JclGetExceptStackListToStrings(ThreadID: DWORD; Strings: TStrings;
  IncludeModuleName: Boolean = False; IncludeAddressOffset: Boolean = False;
  IncludeStartProcLineOffset: Boolean = False; IncludeVAddress: Boolean = False): Boolean;

// helper function for DUnit runtime memory leak check
procedure JclClearGlobalStackData;

// Exception frame info routines
type
  PJmpInstruction = ^TJmpInstruction;
  TJmpInstruction = packed record // from System.pas
    OpCode: Byte;
    Distance: Longint;
  end;

  TExcDescEntry = record // from System.pas
    VTable: Pointer;
    Handler: Pointer;
  end;

  PExcDesc = ^TExcDesc;
  TExcDesc = packed record // from System.pas
    JMP: TJmpInstruction;
    case Integer of
      0:
        (Instructions: array [0..0] of Byte);
      1:
       (Cnt: Integer;
        ExcTab: array [0..0] of TExcDescEntry);
  end;

  PExcFrame = ^TExcFrame;
  TExcFrame =  record // from System.pas
    Next: PExcFrame;
    Desc: PExcDesc;
    FramePointer: Pointer;
    case Integer of
      0:
        ();
      1:
        (ConstructedObject: Pointer);
      2:
        (SelfOfMethod: Pointer);
  end;

  PJmpTable = ^TJmpTable;
  TJmpTable = packed record
    OPCode: Word; // FF 25 = JMP DWORD PTR [$xxxxxxxx], encoded as $25FF
    Ptr: Pointer;
  end;

  TExceptFrameKind =
    (efkUnknown, efkFinally, efkAnyException, efkOnException, efkAutoException);

  TJclExceptFrame = class(TObject)
  private
    FFrameKind: TExceptFrameKind;
    FFrameLocation: Pointer;
    FCodeLocation: Pointer;
    FExcTab: array of TExcDescEntry;
  protected
    procedure AnalyseExceptFrame(AExcDesc: PExcDesc);
  public
    constructor Create(AFrameLocation: Pointer; AExcDesc: PExcDesc);
    function Handles(ExceptObj: TObject): Boolean;
    function HandlerInfo(ExceptObj: TObject; out HandlerAt: Pointer): Boolean;
    property CodeLocation: Pointer read FCodeLocation;
    property FrameLocation: Pointer read FFrameLocation;
    property FrameKind: TExceptFrameKind read FFrameKind;
  end;

  TJclExceptFrameList = class(TJclStackBaseList)
  private
    FIgnoreLevels: Integer;
    function GetItems(Index: Integer): TJclExceptFrame;
  protected
    function AddFrame(AFrame: PExcFrame): TJclExceptFrame;
  public
    constructor Create(AIgnoreLevels: Integer);
    procedure TraceExceptionFrames;
    property Items[Index: Integer]: TJclExceptFrame read GetItems;
    property IgnoreLevels: Integer read FIgnoreLevels write FIgnoreLevels;
  end;

const
  EnvironmentVarNtSymbolPath = '_NT_SYMBOL_PATH';                    // do not localize
  EnvironmentVarAlternateNtSymbolPath = '_NT_ALTERNATE_SYMBOL_PATH'; // do not localize
  MaxStackTraceItems = 4096;

// JCL binary debug data generator and scanner
const
  JclDbgDataSignature = $4742444A; // JDBG
  JclDbgDataResName   = AnsiString('JCLDEBUG'); // do not localize
  JclDbgHeaderVersion = 1; // JCL 1.11 and 1.20

  JclDbgFileExtension = '.jdbg'; // do not localize
  JclMapFileExtension = '.map';  // do not localize
  DrcFileExtension = '.drc';  // do not localize

// Global exceptional stack tracker enable routines and variables
type
  TJclStackTrackingOption =
    (stStack, stExceptFrame, stRawMode, stAllModules, stStaticModuleList,
     stDelayedTrace, stTraceAllExceptions, stMainThreadOnly, stDisableIfDebuggerAttached);
  TJclStackTrackingOptions = set of TJclStackTrackingOption;

var
  JclStackTrackingOptions: TJclStackTrackingOptions = [stStack];

  { JclDebugInfoSymbolPaths specifies a list of paths, separated by ';', in
    which the DebugInfoSymbol scanner should look for symbol information. }
  JclDebugInfoSymbolPaths: string = '';


function LoadedModulesList(const List: TStrings; ProcessID: DWORD; HandlesOnly: Boolean = False): Boolean;

function ModuleFromAddr(const Addr: Pointer): HMODULE;
function IsSystemModule(const Module: HMODULE): Boolean;


function GetModulePath(const Module: HMODULE): string;


// #######################################################################################################################################################
// #######################################################################################################################################################
// #######################################################################################################################################################

implementation

uses
  System.RTLConsts,
  System.Types, // for inlining TList.Remove
  System.Generics.Collections;

function LoadedModulesList(const List: TStrings; ProcessID: DWORD; HandlesOnly: Boolean): Boolean;

  procedure AddToList(ProcessHandle: THandle; Module: HMODULE);
  var
    FileName: array [0..MAX_PATH] of Char;
    ModuleInfo: TModuleInfo;
  begin
    ModuleInfo.EntryPoint := nil;
    if GetModuleInformation(ProcessHandle, Module, @ModuleInfo, SizeOf(ModuleInfo)) then
    begin
      if HandlesOnly then
        List.AddObject('', Pointer(ModuleInfo.lpBaseOfDll))
      else
      if GetModuleFileNameEx(ProcessHandle, Module, Filename, SizeOf(Filename)) > 0 then
        List.AddObject(FileName, Pointer(ModuleInfo.lpBaseOfDll));
    end;
  end;

  function EnumModulesVQ(ProcessHandle: THandle): Boolean;
  var
    MemInfo: TMemoryBasicInformation;
    Base: PChar;
    LastAllocBase: Pointer;
    Res: DWORD;
  begin
    Base := nil;
    LastAllocBase := nil;
    FillChar(MemInfo, SizeOf(MemInfo), #0);
    Res := VirtualQueryEx(ProcessHandle, Base, MemInfo, SizeOf(MemInfo));
    Result := (Res = SizeOf(MemInfo));
    while Res = SizeOf(MemInfo) do
    begin
      if MemInfo.AllocationBase <> LastAllocBase then
      begin
        if MemInfo.Type_9 = MEM_IMAGE then
          AddToList(ProcessHandle, HMODULE(MemInfo.AllocationBase));
        LastAllocBase := MemInfo.AllocationBase;
      end;
      Inc(Base, MemInfo.RegionSize);
      Res := VirtualQueryEx(ProcessHandle, Base, MemInfo, SizeOf(MemInfo));
    end;
  end;

  function EnumModulesPS: Boolean;
  var
    ProcessHandle: THandle;
    Needed: DWORD;
    Modules: array of THandle;
    I, Cnt: Integer;
  begin
    Result := False;
    ProcessHandle := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, ProcessID);
    if ProcessHandle <> 0 then
    try
      Needed := 0;
      Result := EnumProcessModules(ProcessHandle, nil, 0, Needed);
      if Result then
      begin
        Cnt := Needed div SizeOf(HMODULE);
        SetLength(Modules, Cnt);
        if EnumProcessModules(ProcessHandle, @Modules[0], Needed, Needed) then
          for I := 0 to Cnt - 1 do
            AddToList(ProcessHandle, Modules[I]);
      end
      else
        Result := EnumModulesVQ(ProcessHandle);
    finally
      CloseHandle(ProcessHandle);
    end;
  end;

begin
  List.BeginUpdate;
  try
    Result := EnumModulesPS
  finally
    List.EndUpdate;
  end;
end;

function IsSystemModule(const Module: HMODULE): Boolean;
var
  CurModule: PLibModule;
begin
  Result := False;
  if Module <> 0 then
  begin
    CurModule := LibModuleList;
    while CurModule <> nil do
    begin
      if CurModule.Instance = Module then
      begin
        Result := True;
        Break;
      end;
      CurModule := CurModule.Next;
    end;
  end;
end;


function JclCheckWinVersion(Major, Minor: Integer): Boolean;
begin
  Result := CheckWin32Version(Major, Minor);
end;


function GetModulePath(const Module: HMODULE): string;
var
  L: Integer;
begin
  L := MAX_PATH + 1;
  SetLength(Result, L);
  L := Winapi.Windows.GetModuleFileName(Module, Pointer(Result), L);
  SetLength(Result, L);
end;


const
  TurboDebuggerSymbolExt = '.tds';

//=== { TJclModuleInfo } =====================================================

constructor TJclTD32ModuleInfo.Create(pModInfo: PModuleInfo);
begin
  Assert(Assigned(pModInfo));
  inherited Create;
  FNameIndex := pModInfo.NameIndex;
  FSegments := @pModInfo.Segments[0];
  FSegmentCount := pModInfo.SegmentCount;
end;

function TJclTD32ModuleInfo.GetSegment(const Idx: Integer): TSegmentInfo;
begin
  Assert((0 <= Idx) and (Idx < FSegmentCount));
  Result := FSegments[Idx];
end;

//=== { TJclLineInfo } =======================================================

constructor TJclTD32LineInfo.Create(ALineNo, AOffset: DWORD);
begin
  inherited Create;
  FLineNo := ALineNo;
  FOffset := AOffset;
end;

//=== { TJclSourceModuleInfo } ===============================================

constructor TJclTD32SourceModuleInfo.Create(pSrcFile: PSourceFileEntry; Base: NativeInt);
type
  PArrayOfWord = ^TArrayOfWord;
  TArrayOfWord = array [0..0] of Word;
var
  I, J: Integer;
  pLineEntry: PLineMappingEntry;
begin
  Assert(Assigned(pSrcFile));
  inherited Create;
  FNameIndex := pSrcFile.NameIndex;
  FLines := TObjectList.Create;
  {$RANGECHECKS OFF}
  for I := 0 to pSrcFile.SegmentCount - 1 do
  begin
    pLineEntry := PLineMappingEntry(Base + pSrcFile.BaseSrcLines[I]);
    for J := 0 to pLineEntry.PairCount - 1 do
      FLines.Add(TJclTD32LineInfo.Create(
        PArrayOfWord(@pLineEntry.Offsets[pLineEntry.PairCount])^[J],
        pLineEntry.Offsets[J]));
  end;

  FSegments := @pSrcFile.BaseSrcLines[pSrcFile.SegmentCount];
  FSegmentCount := pSrcFile.SegmentCount;
  {$IFDEF RANGECHECKS_ON}
  {$RANGECHECKS ON}
  {$ENDIF RANGECHECKS_ON}
end;

destructor TJclTD32SourceModuleInfo.Destroy;
begin
  FreeAndNil(FLines);
  inherited Destroy;
end;

function TJclTD32SourceModuleInfo.GetLine(const Idx: Integer): TJclTD32LineInfo;
begin
  Result := TJclTD32LineInfo(FLines.Items[Idx]);
end;

function TJclTD32SourceModuleInfo.GetLineCount: Integer;
begin
  Result := FLines.Count;
end;

function TJclTD32SourceModuleInfo.GetSegment(const Idx: Integer): TOffsetPair;
begin
  Assert((0 <= Idx) and (Idx < FSegmentCount));
  Result := FSegments[Idx];
end;

function TJclTD32SourceModuleInfo.FindLine(const AAddr: DWORD; out ALine: TJclTD32LineInfo): Boolean;
var
  I: Integer;
begin
  for I := 0 to LineCount - 1 do
    with Line[I] do
    begin
      if AAddr = Offset then
      begin
        Result := True;
        ALine := Line[I];
        Exit;
      end
      else
      if (I > 1) and (Line[I - 1].Offset < AAddr) and (AAddr < Offset) then
      begin
        Result := True;
        ALine := Line[I-1];
        Exit;
      end;
    end;
  Result := False;
  ALine := nil;
end;

//=== { TJclSymbolInfo } =====================================================

constructor TJclTD32SymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create;
  FSymbolType := pSymInfo.SymbolType;
end;

//=== { TJclProcSymbolInfo } =================================================

constructor TJclTD32ProcSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := Proc.NameIndex;
    FOffset := Proc.Offset;
    FSize := Proc.Size;
  end;
end;

//=== { TJclObjNameSymbolInfo } ==============================================

constructor TJclTD32ObjNameSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := ObjName.NameIndex;
    FSignature := ObjName.Signature;
  end;
end;

//=== { TJclDataSymbolInfo } =================================================

constructor TJclTD32DataSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FTypeIndex := Data.TypeIndex;
    FNameIndex := Data.NameIndex;
    FOffset := Data.Offset;
  end;
end;

//=== { TJclWithSymbolInfo } =================================================

constructor TJclTD32WithSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := With32.NameIndex;
    FOffset := With32.Offset;
    FSize := With32.Size;
  end;
end;

//=== { TJclLabelSymbolInfo } ================================================

constructor TJclTD32LabelSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := Label32.NameIndex;
    FOffset := Label32.Offset;
  end;
end;

//=== { TJclConstantSymbolInfo } =============================================

constructor TJclTD32ConstantSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := Constant.NameIndex;
    FTypeIndex := Constant.TypeIndex;
    FValue := Constant.Value;
  end;
end;

//=== { TJclUdtSymbolInfo } ==================================================

constructor TJclTD32UdtSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FNameIndex := Udt.NameIndex;
    FTypeIndex := Udt.TypeIndex;
    FProperties := Udt.Properties;
  end;
end;

//=== { TJclVftPathSymbolInfo } ==============================================

constructor TJclTD32VftPathSymbolInfo.Create(pSymInfo: PSymbolInfo);
begin
  Assert(Assigned(pSymInfo));
  inherited Create(pSymInfo);
  with pSymInfo^ do
  begin
    FRootIndex := VftPath.RootIndex;
    FPathIndex := VftPath.PathIndex;
    FOffset := VftPath.Offset;
  end;
end;

//=== { TJclTD32InfoParser } =================================================

constructor TJclTD32InfoParser.Create(const ATD32Data: TCustomMemoryStream);
begin
  Assert(Assigned(ATD32Data));
  inherited Create;
  FNames := TList.Create;
  FModules := TObjectList.Create;
  FSourceModules := TObjectList.Create;
  FSymbols := TObjectList.Create;
  FProcSymbols := TList.Create;
  FNames.Add(nil);
  FData := ATD32Data;
  FBase := FData.Memory;
  FValidData := IsTD32DebugInfoValid(FBase, FData.Size);
  if FValidData then
    Analyse;
end;

destructor TJclTD32InfoParser.Destroy;
begin
  FreeAndNil(FProcSymbols);
  FreeAndNil(FSymbols);
  FreeAndNil(FSourceModules);
  FreeAndNil(FModules);
  FreeAndNil(FNames);
  inherited Destroy;
end;

procedure TJclTD32InfoParser.Analyse;
var
  I: Integer;
  pDirHeader: PDirectoryHeader;
  pSubsection: Pointer;
begin
  pDirHeader := PDirectoryHeader(LfaToVa(PJclTD32FileSignature(LfaToVa(0)).Offset));
  while True do
  begin
    Assert(pDirHeader.DirEntrySize = SizeOf(TDirectoryEntry));
    {$RANGECHECKS OFF}
    for I := 0 to pDirHeader.DirEntryCount - 1 do
      with pDirHeader.DirEntries[I] do
      begin
        pSubsection := LfaToVa(Offset);
        case SubsectionType of
          SUBSECTION_TYPE_MODULE:
            AnalyseModules(pSubsection, Size);
          SUBSECTION_TYPE_ALIGN_SYMBOLS:
            AnalyseAlignSymbols(pSubsection, Size);
          SUBSECTION_TYPE_SOURCE_MODULE:
            AnalyseSourceModules(pSubsection, Size);
          SUBSECTION_TYPE_NAMES:
            AnalyseNames(pSubsection, Size);
          SUBSECTION_TYPE_GLOBAL_TYPES:
            AnalyseGlobalTypes(pSubsection, Size);
        else
          AnalyseUnknownSubSection(pSubsection, Size);
        end;
      end;
    {$IFDEF RANGECHECKS_ON}
    {$RANGECHECKS ON}
    {$ENDIF RANGECHECKS_ON}
    if pDirHeader.lfoNextDir <> 0 then
      pDirHeader := PDirectoryHeader(LfaToVa(pDirHeader.lfoNextDir))
    else
      Break;
  end;
end;

procedure TJclTD32InfoParser.AnalyseNames(const pSubsection: Pointer; const Size: DWORD);
var
  I, Count, Len: Integer;
  pszName: PAnsiChar;
begin
  Count := PDWORD(pSubsection)^;
  pszName := PAnsiChar(NativeInt(pSubsection) + SizeOf(DWORD));
  if Count > 0 then
  begin
    FNames.Capacity := FNames.Capacity + Count;
    for I := 0 to Count - 1 do
    begin
      // Get the length of the name
      Len := Ord(pszName^);
      Inc(pszName);
      // Get the name
      FNames.Add(pszName);
      // first, skip the length of name
      Inc(pszName, Len);
      // the length is only correct modulo 256 because it is stored on a single byte,
      // so we have to iterate until we find the real end of the string
      while PszName^ <> #0 do
        Inc(pszName, 256);
      // then, skip a NULL at the end
      Inc(pszName, 1);
    end;
  end;
end;



type
  PSymbolTypeInfo = ^TSymbolTypeInfo;
  TSymbolTypeInfo = packed record
    TypeId: DWORD;
    NameIndex: DWORD;  // 0 if unnamed
    Size: Word;        //  size in bytes of the object
    MaxSize: Byte;
    ParentIndex: DWORD;
  end;


procedure TJclTD32InfoParser.AnalyseGlobalTypes(const pTypes: Pointer; const Size: DWORD);
var
  pTyp: PSymbolTypeInfo;
begin
  pTyp := PSymbolTypeInfo(pTypes);
  repeat
    {case pTyp.TypeId of
      TID_VOID: ;
    end;}
    pTyp := PSymbolTypeInfo(NativeInt(pTyp) + pTyp.Size + SizeOf(pTyp^));
  until NativeInt(pTyp) >= NativeInt(pTypes) + Size;
end;

procedure TJclTD32InfoParser.AnalyseAlignSymbols(pSymbols: PSymbolInfos; const Size: DWORD);
var
  Offset: NativeInt;
  pInfo: PSymbolInfo;
  Symbol: TJclTD32SymbolInfo;
begin
  Offset := NativeInt(@pSymbols.Symbols[0]) - NativeInt(pSymbols);
  while Offset < Size do
  begin
    pInfo := PSymbolInfo(NativeInt(pSymbols) + Offset);
    case pInfo.SymbolType of
      SYMBOL_TYPE_LPROC32:
        begin
          Symbol := TJclTD32LocalProcSymbolInfo.Create(pInfo);
          FProcSymbols.Add(Symbol);
        end;
      SYMBOL_TYPE_GPROC32:
        begin
          Symbol := TJclTD32GlobalProcSymbolInfo.Create(pInfo);
          FProcSymbols.Add(Symbol);
        end;
      SYMBOL_TYPE_OBJNAME:
        Symbol := TJclTD32ObjNameSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_LDATA32:
        Symbol := TJclTD32LDataSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_GDATA32:
        Symbol := TJclTD32GDataSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_PUB32:
        Symbol := TJclTD32PublicSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_WITH32:
        Symbol := TJclTD32WithSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_LABEL32:
        Symbol := TJclTD32LabelSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_CONST:
        Symbol := TJclTD32ConstantSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_UDT:
        Symbol := TJclTD32UdtSymbolInfo.Create(pInfo);
      SYMBOL_TYPE_VFTPATH32:
        Symbol := TJclTD32VftPathSymbolInfo.Create(pInfo);
    else
      Symbol := nil;
    end;
    if Assigned(Symbol) then
      FSymbols.Add(Symbol);
    Inc(Offset, pInfo.Size + SizeOf(pInfo.Size));
  end;
end;

procedure TJclTD32InfoParser.AnalyseModules(pModInfo: PModuleInfo; const Size: DWORD);
begin
  FModules.Add(TJclTD32ModuleInfo.Create(pModInfo));
end;

procedure TJclTD32InfoParser.AnalyseSourceModules(pSrcModInfo: PSourceModuleInfo; const Size: DWORD);
var
  I: Integer;
  pSrcFile: PSourceFileEntry;
begin
  {$RANGECHECKS OFF}
  for I := 0 to pSrcModInfo.FileCount - 1 do
  begin
    pSrcFile := PSourceFileEntry(NativeInt(pSrcModInfo) + pSrcModInfo.BaseSrcFiles[I]);
    if pSrcFile.NameIndex > 0 then
      FSourceModules.Add(TJclTD32SourceModuleInfo.Create(pSrcFile, NativeInt(pSrcModInfo)));
  end;
  {$IFDEF RANGECHECKS_ON}
  {$RANGECHECKS ON}
  {$ENDIF RANGECHECKS_ON}
end;

procedure TJclTD32InfoParser.AnalyseUnknownSubSection(const pSubsection: Pointer; const Size: DWORD);
begin
  // do nothing
end;

function TJclTD32InfoParser.GetModule(const Idx: Integer): TJclTD32ModuleInfo;
begin
  Result := TJclTD32ModuleInfo(FModules.Items[Idx]);
end;

function TJclTD32InfoParser.GetModuleCount: Integer;
begin
  Result := FModules.Count;
end;

function TJclTD32InfoParser.GetName(const Idx: Integer): string;
begin
  Result := UTF8ToString(PAnsiChar(FNames.Items[Idx]));
end;

function TJclTD32InfoParser.GetNameCount: Integer;
begin
  Result := FNames.Count;
end;

function TJclTD32InfoParser.GetSourceModule(const Idx: Integer): TJclTD32SourceModuleInfo;
begin
  Result := TJclTD32SourceModuleInfo(FSourceModules.Items[Idx]);
end;

function TJclTD32InfoParser.GetSourceModuleCount: Integer;
begin
  Result := FSourceModules.Count;
end;

function TJclTD32InfoParser.GetSymbol(const Idx: Integer): TJclTD32SymbolInfo;
begin
  Result := TJclTD32SymbolInfo(FSymbols.Items[Idx]);
end;

function TJclTD32InfoParser.GetSymbolCount: Integer;
begin
  Result := FSymbols.Count;
end;

function TJclTD32InfoParser.GetProcSymbol(const Idx: Integer): TJclTD32ProcSymbolInfo;
begin
  Result := TJclTD32ProcSymbolInfo(FProcSymbols.Items[Idx]);
end;

function TJclTD32InfoParser.GetProcSymbolCount: Integer;
begin
  Result := FProcSymbols.Count;
end;

function TJclTD32InfoParser.FindModule(const AAddr: DWORD; out AMod: TJclTD32ModuleInfo): Boolean;
var
  I, J: Integer;
begin
  if ValidData then
    for I := 0 to ModuleCount - 1 do
    with Modules[I] do
      for J := 0 to SegmentCount - 1 do
      begin
        if (FSegments[J].Flags = 1) and (AAddr >= FSegments[J].Offset) and (AAddr - FSegments[J].Offset <= Segment[J].Size) then
        begin
          Result := True;
          AMod := Modules[I];
          Exit;
        end;
      end;
  Result := False;
  AMod := nil;
end;

function TJclTD32InfoParser.FindSourceModule(const AAddr: DWORD; out ASrcMod: TJclTD32SourceModuleInfo): Boolean;
var
  I, J: Integer;
begin
  if ValidData then
    for I := 0 to SourceModuleCount - 1 do
    with SourceModules[I] do
      for J := 0 to SegmentCount - 1 do
        with Segment[J] do
          if (StartOffset <= AAddr) and (AAddr < EndOffset) then
          begin
            Result := True;
            ASrcMod := SourceModules[I];
            Exit;
          end;
  ASrcMod := nil;
  Result := False;
end;

function TJclTD32InfoParser.FindProc(const AAddr: DWORD; out AProc: TJclTD32ProcSymbolInfo): Boolean;
var
  I: Integer;
begin
  if ValidData then
    for I := 0 to ProcSymbolCount - 1 do
    begin
      AProc := ProcSymbols[I];
      with AProc do
        if (Offset <= AAddr) and (AAddr < Offset + Size) then
        begin
          Result := True;
          Exit;
        end;
    end;
  AProc := nil;
  Result := False;
end;

class function TJclTD32InfoParser.IsTD32DebugInfoValid(
  const DebugData: Pointer; const DebugDataSize: LongWord): Boolean;
var
  Sign: TJclTD32FileSignature;
  EndOfDebugData: NativeInt;
begin
  Assert(not IsBadReadPtr(DebugData, DebugDataSize));
  Result := False;
  EndOfDebugData := NativeInt(DebugData) + DebugDataSize;
  if DebugDataSize > SizeOf(Sign) then
  begin
    Sign := PJclTD32FileSignature(EndOfDebugData - SizeOf(Sign))^;
    if IsTD32Sign(Sign) and (Sign.Offset <= DebugDataSize) then
    begin
      Sign := PJclTD32FileSignature(EndOfDebugData - Sign.Offset)^;
      Result := IsTD32Sign(Sign);
    end;
  end;
end;

class function TJclTD32InfoParser.IsTD32Sign(const Sign: TJclTD32FileSignature): Boolean;
begin
  Result := (Sign.Signature = Borland32BitSymbolFileSignatureForDelphi) or
    (Sign.Signature = Borland32BitSymbolFileSignatureForBCB);
end;

function TJclTD32InfoParser.LfaToVa(Lfa: DWORD): Pointer;
begin
  Result := Pointer(NativeInt(FBase) + Lfa)
end;

//=== { TJclTD32InfoScanner } ================================================

function TJclTD32InfoScanner.LineNumberFromAddr(AAddr: DWORD): Integer;
var
  Dummy: Integer;
begin
  Result := LineNumberFromAddr(AAddr, Dummy);
end;

function TJclTD32InfoScanner.LineNumberFromAddr(AAddr: DWORD; out Offset: Integer): Integer;
var
  ASrcMod: TJclTD32SourceModuleInfo;
  ALine: TJclTD32LineInfo;
begin
  if FindSourceModule(AAddr, ASrcMod) and ASrcMod.FindLine(AAddr, ALine) then
  begin
    Result := ALine.LineNo;
    Offset := AAddr - ALine.Offset;
  end
  else
  begin
    Result := 0;
    Offset := 0;
  end;
end;

function TJclTD32InfoScanner.ModuleNameFromAddr(AAddr: DWORD): string;
var
  AMod: TJclTD32ModuleInfo;
begin
  if FindModule(AAddr, AMod) then
    Result := Names[AMod.NameIndex]
  else
    Result := '';
end;

function TJclTD32InfoScanner.ProcNameFromAddr(AAddr: DWORD): string;
var
  Dummy: Integer;
begin
  Result := ProcNameFromAddr(AAddr, Dummy);
end;

function TJclTD32InfoScanner.ProcNameFromAddr(AAddr: DWORD; out Offset: Integer): string;
var
  AProc: TJclTD32ProcSymbolInfo;

  function FormatProcName(const ProcName: string): string;
  var
    pchSecondAt, P: PChar;
  begin
    Result := ProcName;
    if (Length(ProcName) > 0) and (ProcName[1] = '@') then
    begin
      pchSecondAt := StrScan(PChar(Copy(ProcName, 2, Length(ProcName) - 1)), '@');
      if pchSecondAt <> nil then
      begin
        Inc(pchSecondAt);
        Result := pchSecondAt;
        P := PChar(Result);
        while P^ <> #0 do
        begin
          if (pchSecondAt^ = '@') and ((pchSecondAt - 1)^ <> '@') then
            P^ := '.';
          Inc(P);
          Inc(pchSecondAt);
        end;
      end;
    end;
  end;

begin
  if FindProc(AAddr, AProc) then
  begin
    Result := FormatProcName(Names[AProc.NameIndex]);
    Offset := AAddr - AProc.Offset;
  end
  else
  begin
    Result := '';
    Offset := 0;
  end;
end;

function TJclTD32InfoScanner.SourceNameFromAddr(AAddr: DWORD): string;
var
  ASrcMod: TJclTD32SourceModuleInfo;
begin
  if FindSourceModule(AAddr, ASrcMod) then
    Result := Names[ASrcMod.NameIndex];
end;


//=== { TJclPeBorTD32Image } =================================================

procedure TJclPeBorTD32Image.AfterOpen;
begin
  inherited AfterOpen;
  CheckDebugData;
end;

procedure TJclPeBorTD32Image.CheckDebugData;
begin
  FIsTD32DebugPresent := IsDebugInfoInImage(FTD32DebugData);
  if not FIsTD32DebugPresent then
    FIsTD32DebugPresent := IsDebugInfoInTds(FTD32DebugData);
  if FIsTD32DebugPresent then
  begin
    FTD32Scanner := TJclTD32InfoScanner.Create(FTD32DebugData);
    if not FTD32Scanner.ValidData then
    begin
      ClearDebugData;
      if not NoExceptions then
        raise Exception.CreateFmt('File [%s] has not TD32 debug information!', [FileName]);
    end;
  end;
end;

procedure TJclPeBorTD32Image.Clear;
begin
  ClearDebugData;
  inherited Clear;
end;

procedure TJclPeBorTD32Image.ClearDebugData;
begin
  FIsTD32DebugPresent := False;
  FreeAndNil(FTD32Scanner);
  FreeAndNil(FTD32DebugData);
end;

function TJclPeBorTD32Image.IsDebugInfoInImage(var DataStream: TCustomMemoryStream): Boolean;
var
  DebugDir: TImageDebugDirectory;
  BugDataStart: Pointer;
  DebugDataSize: Integer;
begin
  Result := False;
  DataStream := nil;
  if IsBorlandImage and (DebugList.Count = 1) then
  begin
    DebugDir := DebugList[0];
    if DebugDir._Type = IMAGE_DEBUG_TYPE_UNKNOWN then
    begin
      BugDataStart := RvaToVa(DebugDir.AddressOfRawData);
      DebugDataSize := DebugDir.SizeOfData;
      Result := TJclTD32InfoParser.IsTD32DebugInfoValid(BugDataStart, DebugDataSize);
      if Result then
        DataStream := TJclReferenceMemoryStream.Create(BugDataStart, DebugDataSize);
    end;
  end;
end;

function TJclPeBorTD32Image.IsDebugInfoInTds(var DataStream: TCustomMemoryStream): Boolean;
begin
  Result := False;
end;

//=== { TJclReferenceMemoryStream } ==========================================

constructor TJclReferenceMemoryStream.Create(const Ptr: Pointer; Size: Longint);
begin
  Assert(not IsBadReadPtr(Ptr, Size));
  inherited Create;
  SetPointer(Ptr, Size);
end;

function TJclReferenceMemoryStream.Write(const Buffer; Count: Longint): Longint;
begin
  raise Exception.Create('Can not write to a read-only memory stream');
end;

const
  MANIFESTExtension = '.manifest';

  DebugSectionName    = '.debug';
  ReadOnlySectionName = '.rdata';

  BinaryExtensionLibrary = '.dll';

  CompilerExtensionDCP   = '.dcp';
  BinaryExtensionPackage = '.bpl';

  PackageInfoResName    = 'PACKAGEINFO';
  DescriptionResName    = 'DESCRIPTION';
  PackageOptionsResName = 'PACKAGEOPTIONS';
  DVclAlResName         = 'DVCLAL';

// Helper routines
function AddFlagTextRes(var Text: string; const FlagText: PResStringRec; const Value, Mask: Cardinal): Boolean;
begin
  Result := (Value and Mask <> 0);
  if Result then
  begin
    if Length(Text) > 0 then
      Text := Text + ', ';
    Text := Text + LoadResString(FlagText);
  end;
end;

function CompareResourceName(T1, T2: PChar): Boolean;
var
  Long1, Long2: LongRec;
begin
  {$IFDEF WIN64}
  Long1 := LongRec(Int64Rec(T1).Lo);
  Long2 := LongRec(Int64Rec(T2).Lo);
  if (Int64Rec(T1).Hi = 0) and (Int64Rec(T2).Hi = 0) and (Long1.Hi = 0) and (Long2.Hi = 0) then
  {$ENDIF WIN64}
  {$IFDEF WIN32}
  Long1 := LongRec(T1);
  Long2 := LongRec(T2);
  if (Long1.Hi = 0) or (Long2.Hi = 0) then
  {$ENDIF WIN32}
    Result := Long1.Lo = Long2.Lo
  else
    Result := (StrIComp(T1, T2) = 0);
end;


function InternalImportedLibraries(const FileName: TFileName;
  Recursive, FullPathName: Boolean; ExternalCache: TJclPeImagesCache): TStringList;
var
  Cache: TJclPeImagesCache;

  procedure ProcessLibraries(const AFileName: TFileName);
  var
    I: Integer;
    S: TFileName;
    ImportLib: TJclPeImportLibItem;
  begin
    with Cache[AFileName].ImportList do
      for I := 0 to Count - 1 do
      begin
        ImportLib := Items[I];
        if FullPathName then
          S := ImportLib.FileName
        else
          S := TFileName(ImportLib.Name);
        if Result.IndexOf(S) = -1 then
        begin
          Result.Add(S);
          if Recursive then
            ProcessLibraries(ImportLib.FileName);
        end;
      end;
  end;

begin
  if ExternalCache = nil then
    Cache := TJclPeImagesCache.Create
  else
    Cache := ExternalCache;
  try
    Result := TStringList.Create;
    try
      Result.Sorted := True;
      Result.Duplicates := dupIgnore;
      ProcessLibraries(FileName);
    except
      FreeAndNil(Result);
      raise;
    end;
  finally
    if ExternalCache = nil then
      Cache.Free;
  end;
end;

function CharIsValidIdentifierLetter(const C: Char): Boolean;
begin
  case C of
    '0'..'9', 'A'..'Z', 'a'..'z', '_':
      Result := True;
  else
    Result := False;
  end;
end;

// Smart name compare function
function PeStripFunctionAW(const FunctionName: string): string;
var
  L: Integer;
begin
  Result := FunctionName;
  L := Length(Result);
  if (L > 1) then
    case Result[L] of
      'A', 'W':
        if CharIsValidIdentifierLetter(Result[L - 1]) then
          Delete(Result, L, 1);
    end;
end;

function PeSmartFunctionNameSame(const ComparedName, FunctionName: string;
  Options: TJclSmartCompOptions): Boolean;
var
  S: string;
begin
  if scIgnoreCase in Options then
    Result := CompareText(FunctionName, ComparedName) = 0
  else
    Result := (FunctionName = ComparedName);
  if (not Result) and not (scSimpleCompare in Options) then
  begin
    if Length(FunctionName) > 0 then
    begin
      S := PeStripFunctionAW(FunctionName);
      if scIgnoreCase in Options then
        Result := CompareText(S, ComparedName) = 0
      else
        Result := (S = ComparedName);
    end
    else
      Result := False;
  end;
end;

//=== { TJclPeImagesCache } ==================================================

constructor TJclPeImagesCache.Create;
begin
  inherited Create;
  FList := TStringList.Create;
  FList.Sorted := True;
  FList.Duplicates := dupIgnore;
end;

destructor TJclPeImagesCache.Destroy;
begin
  Clear;
  FreeAndNil(FList);
  inherited Destroy;
end;

procedure TJclPeImagesCache.Clear;
var
  I: Integer;
begin
  with FList do
    for I := 0 to Count - 1 do
      Objects[I].Free;
  FList.Clear;
end;

function TJclPeImagesCache.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TJclPeImagesCache.GetImages(const FileName: TFileName): TJclPeImage;
var
  I: Integer;
begin
  I := FList.IndexOf(FileName);
  if I = -1 then
  begin
    Result := GetPeImageClass.Create(True);
    Result.FileName := FileName;
    FList.AddObject(FileName, Result);
  end
  else
    Result := TJclPeImage(FList.Objects[I]);
end;

function TJclPeImagesCache.GetPeImageClass: TJclPeImageClass;
begin
  Result := TJclPeImage;
end;

//=== { TJclPeImageBaseList } ================================================

constructor TJclPeImageBaseList.Create(AImage: TJclPeImage);
begin
  inherited Create(True);
  FImage := AImage;
end;

// Import sort functions

function ImportSortByName(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeImportFuncItem(Item1).Name, TJclPeImportFuncItem(Item2).Name);
  if Result = 0 then
    Result := CompareStr(TJclPeImportFuncItem(Item1).ImportLib.Name, TJclPeImportFuncItem(Item2).ImportLib.Name);
  if Result = 0 then
    Result := TJclPeImportFuncItem(Item1).Ordinal - TJclPeImportFuncItem(Item2).Ordinal;
end;

function ImportSortByNameDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ImportSortByName(Item2, Item1);
end;

function ImportSortByHint(Item1, Item2: Pointer): Integer;
begin
  Result := TJclPeImportFuncItem(Item1).Hint - TJclPeImportFuncItem(Item2).Hint;
end;

function ImportSortByHintDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ImportSortByHint(Item2, Item1);
end;

function ImportSortByDll(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeImportFuncItem(Item1).ImportLib.Name,
    TJclPeImportFuncItem(Item2).ImportLib.Name);
  if Result = 0 then
    Result := ImportSortByName(Item1, Item2);
end;

function ImportSortByDllDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ImportSortByDll(Item2, Item1);
end;

function ImportSortByOrdinal(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeImportFuncItem(Item1).ImportLib.Name,
    TJclPeImportFuncItem(Item2).ImportLib.Name);
  if Result = 0 then
    Result := TJclPeImportFuncItem(Item1).Ordinal -  TJclPeImportFuncItem(Item2).Ordinal;
end;

function ImportSortByOrdinalDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ImportSortByOrdinal(Item2, Item1);
end;

function GetImportSortFunction(SortType: TJclPeImportSort; Descending: Boolean): TListSortCompare;
const
  SortFunctions: array [TJclPeImportSort, Boolean] of TListSortCompare =
    ((ImportSortByName, ImportSortByNameDESC),
     (ImportSortByOrdinal, ImportSortByOrdinalDESC),
     (ImportSortByHint, ImportSortByHintDESC),
     (ImportSortByDll, ImportSortByDllDESC)
    );
begin
  Result := SortFunctions[SortType, Descending];
end;

function ImportLibSortByIndex(Item1, Item2: Pointer): Integer;
begin
  Result := TJclPeImportLibItem(Item1).ImportDirectoryIndex -
    TJclPeImportLibItem(Item2).ImportDirectoryIndex;
end;

function ImportLibSortByName(Item1, Item2: Pointer): Integer;
begin
  Result := AnsiCompareStr(TJclPeImportLibItem(Item1).Name, TJclPeImportLibItem(Item2).Name);
  if Result = 0 then
    Result := ImportLibSortByIndex(Item1, Item2);
end;

function GetImportLibSortFunction(SortType: TJclPeImportLibSort): TListSortCompare;
const
  SortFunctions: array [TJclPeImportLibSort] of TListSortCompare =
    (ImportLibSortByName, ImportLibSortByIndex);
begin
  Result := SortFunctions[SortType];
end;

//=== { TJclPeImportFuncItem } ===============================================

constructor TJclPeImportFuncItem.Create(AImportLib: TJclPeImportLibItem;
  AOrdinal: Word; AHint: Word; const AName: string);
begin
  inherited Create;
  FImportLib := AImportLib;
  FOrdinal := AOrdinal;
  FHint := AHint;
  FName := AName;
  FResolveCheck := icNotChecked;
  FIndirectImportName := False;
end;

function TJclPeImportFuncItem.GetIsByOrdinal: Boolean;
begin
  Result := FOrdinal <> 0;
end;

procedure TJclPeImportFuncItem.SetIndirectImportName(const Value: string);
begin
  FName := Value;
  FIndirectImportName := True;
end;

procedure TJclPeImportFuncItem.SetName(const Value: string);
begin
  FName := Value;
  FIndirectImportName := False;
end;

procedure TJclPeImportFuncItem.SetResolveCheck(Value: TJclPeResolveCheck);
begin
  FResolveCheck := Value;
end;

//=== { TJclPeImportLibItem } ================================================

constructor TJclPeImportLibItem.Create(AImage: TJclPeImage;
  AImportDescriptor: Pointer; AImportKind: TJclPeImportKind; const AName: string;
  AThunk: Pointer);
begin
  inherited Create(AImage);
  FTotalResolveCheck := icNotChecked;
  FImportDescriptor := AImportDescriptor;
  FImportKind := AImportKind;
  FName := AName;
  FThunk := AThunk;
  FThunkData := AThunk;
end;

procedure TJclPeImportLibItem.CheckImports(ExportImage: TJclPeImage);
var
  I: Integer;
  ExportList: TJclPeExportFuncList;
begin
  if ExportImage.StatusOK then
  begin
    FTotalResolveCheck := icResolved;
    ExportList := ExportImage.ExportList;
    for I := 0 to Count - 1 do
    begin
      with Items[I] do
        if IsByOrdinal then
        begin
          if ExportList.OrdinalValid(Ordinal) then
            SetResolveCheck(icResolved)
          else
          begin
            SetResolveCheck(icUnresolved);
            Self.FTotalResolveCheck := icUnresolved;
          end;
        end
        else
        begin
          if ExportList.ItemFromName[Items[I].Name] <> nil then
            SetResolveCheck(icResolved)
          else
          begin
            SetResolveCheck(icUnresolved);
            Self.FTotalResolveCheck := icUnresolved;
          end;
        end;
    end;
  end
  else
  begin
    FTotalResolveCheck := icUnresolved;
    for I := 0 to Count - 1 do
      Items[I].SetResolveCheck(icUnresolved);
  end;
end;

function IMAGE_ORDINAL64(Ordinal: ULONGLONG): ULONGLONG; inline;
begin
  Result := (Ordinal and $FFFF);
end;

function IMAGE_ORDINAL32(Ordinal: DWORD): DWORD; inline;
begin
  Result := (Ordinal and $0000FFFF);
end;

procedure TJclPeImportLibItem.CreateList;
  procedure CreateList32;
  var
    Thunk32: PImageThunkData32;
    OrdinalName: PImageImportByName;
    Ordinal, Hint: Word;
    Name: PAnsiChar;
    ImportName: string;
  begin
    Thunk32 := PImageThunkData32(FThunk);
    while Thunk32^.Function_ <> 0 do
    begin
      Ordinal := 0;
      Hint := 0;
      Name := nil;
      if Thunk32^.Ordinal and IMAGE_ORDINAL_FLAG32 = 0 then
      begin
        case ImportKind of
          ikImport, ikBoundImport:
            begin
              OrdinalName := PImageImportByName(Image.RvaToVa(Thunk32^.AddressOfData));
              Hint := OrdinalName.Hint;
              Name := OrdinalName.Name;
            end;
          ikDelayImport:
            begin
              OrdinalName := PImageImportByName(Image.RvaToVaEx(Thunk32^.AddressOfData));
              Hint := OrdinalName.Hint;
              Name := OrdinalName.Name;
            end;
        end;
      end
      else
        Ordinal := IMAGE_ORDINAL32(Thunk32^.Ordinal);
      ImportName := string(Name);
      Add(TJclPeImportFuncItem.Create(Self, Ordinal, Hint, ImportName));
      Inc(Thunk32);
    end;
  end;

  procedure CreateList64;
  var
    Thunk64: PImageThunkData64;
    OrdinalName: PImageImportByName;
    Ordinal, Hint: Word;
    Name: PAnsiChar;
    ImportName: string;
  begin
    Thunk64 := PImageThunkData64(FThunk);
    while Thunk64^.Function_ <> 0 do
    begin
      Ordinal := 0;
      Hint := 0;
      Name := nil;
      if Thunk64^.Ordinal and IMAGE_ORDINAL_FLAG64 = 0 then
      begin
        case ImportKind of
          ikImport, ikBoundImport:
            begin
              OrdinalName := PImageImportByName(Image.RvaToVa(Thunk64^.AddressOfData));
              Hint := OrdinalName.Hint;
              Name := OrdinalName.Name;
            end;
          ikDelayImport:
            begin
              OrdinalName := PImageImportByName(Image.RvaToVaEx(Thunk64^.AddressOfData));
              Hint := OrdinalName.Hint;
              Name := OrdinalName.Name;
            end;
        end;
      end
      else
        Ordinal := IMAGE_ORDINAL64(Thunk64^.Ordinal);
      ImportName := string(Name);
      Add(TJclPeImportFuncItem.Create(Self, Ordinal, Hint, ImportName));
      Inc(Thunk64);
    end;
  end;
begin
  if FThunk = nil then
    Exit;

  case Image.Target of
    taWin32:
      CreateList32;
    taWin64:
      CreateList64;
  end;

  FThunk := nil;
end;

function TJclPeImportLibItem.GetCount: Integer;
begin
  if FThunk <> nil then
    CreateList;
  Result := inherited Count;
end;

function TJclPeImportLibItem.GetFileName: TFileName;
begin
  Result := Image.ExpandModuleName(Name);
end;

function TJclPeImportLibItem.GetItems(Index: Integer): TJclPeImportFuncItem;
begin
  Result := TJclPeImportFuncItem(Get(Index));
end;

function TJclPeImportLibItem.GetName: string;
begin
  Result := AnsiLowerCase(OriginalName);
end;

function TJclPeImportLibItem.GetThunkData32: PImageThunkData32;
begin
  if Image.Target = taWin32 then
    Result := FThunkData
  else
    Result := nil;
end;

function TJclPeImportLibItem.GetThunkData64: PImageThunkData64;
begin
  if Image.Target = taWin64 then
    Result := FThunkData
  else
    Result := nil;
end;

procedure TJclPeImportLibItem.SetImportDirectoryIndex(Value: Integer);
begin
  FImportDirectoryIndex := Value;
end;

procedure TJclPeImportLibItem.SetImportKind(Value: TJclPeImportKind);
begin
  FImportKind := Value;
end;

procedure TJclPeImportLibItem.SetSorted(Value: Boolean);
begin
  FSorted := Value;
end;

procedure TJclPeImportLibItem.SetThunk(Value: Pointer);
begin
  FThunk := Value;
  FThunkData := Value;
end;

procedure TJclPeImportLibItem.SortList(SortType: TJclPeImportSort; Descending: Boolean);
begin
  if not FSorted or (SortType <> FLastSortType) or (Descending <> FLastSortDescending) then
  begin
    GetCount; // create list if it wasn't created
    Sort(GetImportSortFunction(SortType, Descending));
    FLastSortType := SortType;
    FLastSortDescending := Descending;
    FSorted := True;
  end;
end;

//=== { TJclPeImportList } ===================================================

constructor TJclPeImportList.Create(AImage: TJclPeImage);
begin
  inherited Create(AImage);
  FAllItemsList := TList.Create;
  FAllItemsList.Capacity := 256;
  FUniqueNamesList := TStringList.Create;
  FUniqueNamesList.Sorted := True;
  FUniqueNamesList.Duplicates := dupIgnore;
  FLastAllSortType := isName;
  FLastAllSortDescending := False;
  CreateList;
end;

destructor TJclPeImportList.Destroy;
var
  I: Integer;
begin
  FreeAndNil(FAllItemsList);
  FreeAndNil(FUniqueNamesList);
  for I := 0 to Length(FparallelImportTable) - 1 do
    FreeMem(FparallelImportTable[I]);
  inherited Destroy;
end;

procedure TJclPeImportList.CheckImports(PeImageCache: TJclPeImagesCache);
var
  I: Integer;
  ExportPeImage: TJclPeImage;
begin
  Image.CheckNotAttached;
  if PeImageCache <> nil then
    ExportPeImage := nil // to make the compiler happy
  else
    ExportPeImage := TJclPeImage.Create(True);
  try
    for I := 0 to Count - 1 do
      if Items[I].TotalResolveCheck = icNotChecked then
      begin
        if PeImageCache <> nil then
          ExportPeImage := PeImageCache[Items[I].FileName]
        else
          ExportPeImage.FileName := Items[I].FileName;
        ExportPeImage.ExportList.PrepareForFastNameSearch;
        Items[I].CheckImports(ExportPeImage);
      end;
  finally
    if PeImageCache = nil then
      ExportPeImage.Free;
  end;
end;

procedure TJclPeImportList.CreateList;
  procedure CreateDelayImportList32(DelayImportDesc: PImgDelayDescrV1);
  var
    LibItem: TJclPeImportLibItem;
    UTF8Name: AnsiString;
    LibName: string;
  begin
    while DelayImportDesc^.szName <> nil do
    begin
      UTF8Name := PAnsiChar(Image.RvaToVaEx(DWORD(DelayImportDesc^.szName)));
      LibName := string(UTF8Name);
      LibItem := TJclPeImportLibItem.Create(Image, DelayImportDesc, ikDelayImport,
        LibName, Image.RvaToVaEx(DWORD(DelayImportDesc^.pINT)));
      Add(LibItem);
      FUniqueNamesList.AddObject(AnsiLowerCase(LibItem.Name), LibItem);
      Inc(DelayImportDesc);
    end;
  end;

  procedure CreateDelayImportList64(DelayImportDesc: PImgDelayDescrV2);
  var
    LibItem: TJclPeImportLibItem;
    UTF8Name: AnsiString;
    LibName: string;
  begin
    while DelayImportDesc^.rvaDLLName <> 0 do
    begin
      UTF8Name := PAnsiChar(Image.RvaToVa(DelayImportDesc^.rvaDLLName));
      LibName := string(UTF8Name);
      LibItem := TJclPeImportLibItem.Create(Image, DelayImportDesc, ikDelayImport,
        LibName, Image.RvaToVa(DelayImportDesc^.rvaINT));
      Add(LibItem);
      FUniqueNamesList.AddObject(AnsiLowerCase(LibItem.Name), LibItem);
      Inc(DelayImportDesc);
    end;
  end;
var
  ImportDesc: PImageImportDescriptor;
  LibItem: TJclPeImportLibItem;
  UTF8Name: AnsiString;
  LibName, ModuleName: string;
  DelayImportDesc: Pointer;
  BoundImports, BoundImport: PImageBoundImportDescriptor;
  S: string;
  I: Integer;
  Thunk: Pointer;
begin
  SetCapacity(100);
  with Image do
  begin
    if not StatusOK then
      Exit;
    ImportDesc := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_IMPORT);
    if ImportDesc <> nil then
      while ImportDesc^.Name <> 0 do
      begin
        if ImportDesc^.Union.Characteristics = 0 then
        begin
          if AttachedImage then  // Borland images doesn't have two parallel arrays
            Thunk := nil // see MakeBorlandImportTableForMappedImage method
          else
            Thunk := RvaToVa(ImportDesc^.FirstThunk);
          FLinkerProducer := lrBorland;
        end
        else
        begin
          Thunk := RvaToVa(ImportDesc^.Union.Characteristics);
          FLinkerProducer := lrMicrosoft;
        end;
        UTF8Name := PAnsiChar(RvaToVa(ImportDesc^.Name));
        LibName := string(UTF8Name);
        LibItem := TJclPeImportLibItem.Create(Image, ImportDesc, ikImport, LibName, Thunk);
        Add(LibItem);
        FUniqueNamesList.AddObject(AnsiLowerCase(LibItem.Name), LibItem);
        Inc(ImportDesc);
      end;
    DelayImportDesc := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT);
    if DelayImportDesc <> nil then
    begin
      case Target of
        taWin32:
          CreateDelayImportList32(DelayImportDesc);
        taWin64:
          CreateDelayImportList64(DelayImportDesc);
      end;
    end;
    BoundImports := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT);
    if BoundImports <> nil then
    begin
      BoundImport := BoundImports;
      while BoundImport^.OffsetModuleName <> 0 do
      begin
        UTF8Name := PAnsiChar(NativeInt(BoundImports) + BoundImport^.OffsetModuleName);
        ModuleName := string(UTF8Name);
        S := AnsiLowerCase(ModuleName);
        I := FUniqueNamesList.IndexOf(S);
        if I >= 0 then
          TJclPeImportLibItem(FUniqueNamesList.Objects[I]).SetImportKind(ikBoundImport);
        for I := 1 to BoundImport^.NumberOfModuleForwarderRefs do
          Inc(PImageBoundForwarderRef(BoundImport)); // skip forward information
        Inc(BoundImport);
      end;
    end;
  end;
  for I := 0 to Count - 1 do
    Items[I].SetImportDirectoryIndex(I);
end;

function TJclPeImportList.GetAllItemCount: Integer;
begin
  Result := FAllItemsList.Count;
  if Result = 0 then // we haven't created the list yet -> create unsorted list
  begin
    RefreshAllItems;
    Result := FAllItemsList.Count;
  end;
end;

function TJclPeImportList.GetAllItems(Index: Integer): TJclPeImportFuncItem;
begin
  Result := TJclPeImportFuncItem(FAllItemsList[Index]);
end;

function TJclPeImportList.GetItems(Index: Integer): TJclPeImportLibItem;
begin
  Result := TJclPeImportLibItem(Get(Index));
end;

function TJclPeImportList.GetUniqueLibItemCount: Integer;
begin
  Result := FUniqueNamesList.Count;
end;

function TJclPeImportList.GetUniqueLibItemFromName(const Name: string): TJclPeImportLibItem;
var
  I: Integer;
begin
  I := FUniqueNamesList.IndexOf(Name);
  if I = -1 then
    Result := nil
  else
    Result := TJclPeImportLibItem(FUniqueNamesList.Objects[I]);
end;

function TJclPeImportList.GetUniqueLibItems(Index: Integer): TJclPeImportLibItem;
begin
  Result := TJclPeImportLibItem(FUniqueNamesList.Objects[Index]);
end;

function TJclPeImportList.GetUniqueLibNames(Index: Integer): string;
begin
  Result := FUniqueNamesList[Index];
end;

function TJclPeImportList.MakeBorlandImportTableForMappedImage: Boolean;
var
  FileImage: TJclPeImage;
  I, TableSize: Integer;
begin
  if Image.AttachedImage and (LinkerProducer = lrBorland) and
    (Length(FParallelImportTable) = 0) then
  begin
    FileImage := TJclPeImage.Create(True);
    try
      FileImage.FileName := Image.FileName;
      Result := FileImage.StatusOK;
      if Result then
      begin
        SetLength(FParallelImportTable, FileImage.ImportList.Count);
        for I := 0 to FileImage.ImportList.Count - 1 do
        begin
          Assert(Items[I].ImportKind = ikImport); // Borland doesn't have Delay load or Bound imports
          TableSize := (FileImage.ImportList[I].Count + 1);
          case Image.Target of
            taWin32:
              begin
                TableSize := TableSize * SizeOf(TImageThunkData32);
                GetMem(FParallelImportTable[I], TableSize);
                System.Move(FileImage.ImportList[I].ThunkData32^, FParallelImportTable[I]^, TableSize);
                Items[I].SetThunk(FParallelImportTable[I]);
              end;
            taWin64:
              begin
                TableSize := TableSize * SizeOf(TImageThunkData64);
                GetMem(FParallelImportTable[I], TableSize);
                System.Move(FileImage.ImportList[I].ThunkData64^, FParallelImportTable[I]^, TableSize);
                Items[I].SetThunk(FParallelImportTable[I]);
              end;
          end;
        end;
      end;
    finally
      FileImage.Free;
    end;
  end
  else
    Result := True;
end;

procedure TJclPeImportList.RefreshAllItems;
var
  L, I: Integer;
  LibItem: TJclPeImportLibItem;
begin
  FAllItemsList.Clear;
  for L := 0 to Count - 1 do
  begin
    LibItem := Items[L];
    if (Length(FFilterModuleName) = 0) or (AnsiCompareText(LibItem.Name, FFilterModuleName) = 0) then
      for I := 0 to LibItem.Count - 1 do
        FAllItemsList.Add(LibItem[I]);
  end;
end;

procedure TJclPeImportList.SetFilterModuleName(const Value: string);
begin
  if (FFilterModuleName <> Value) or (FAllItemsList.Count = 0) then
  begin
    FFilterModuleName := Value;
    RefreshAllItems;
    FAllItemsList.Sort(GetImportSortFunction(FLastAllSortType, FLastAllSortDescending));
  end;
end;

function TJclPeImportList.SmartFindName(const CompareName, LibName: string;
  Options: TJclSmartCompOptions): TJclPeImportFuncItem;
var
  L, I: Integer;
  LibItem: TJclPeImportLibItem;
begin
  Result := nil;
  for L := 0 to Count - 1 do
  begin
    LibItem := Items[L];
    if (Length(LibName) = 0) or (AnsiCompareText(LibItem.Name, LibName) = 0) then
      for I := 0 to LibItem.Count - 1 do
        if PeSmartFunctionNameSame(CompareName, LibItem[I].Name, Options) then
        begin
          Result := LibItem[I];
          Break;
        end;
  end;
end;

procedure TJclPeImportList.SortAllItemsList(SortType: TJclPeImportSort; Descending: Boolean);
begin
  GetAllItemCount; // create list if it wasn't created
  FAllItemsList.Sort(GetImportSortFunction(SortType, Descending));
  FLastAllSortType := SortType;
  FLastAllSortDescending := Descending;
end;

procedure TJclPeImportList.SortList(SortType: TJclPeImportLibSort);
begin
  Sort(GetImportLibSortFunction(SortType));
end;

procedure TJclPeImportList.TryGetNamesForOrdinalImports;
var
  LibNamesList: TStringList;
  L, I: Integer;
  LibPeDump: TJclPeImage;

  procedure TryGetNames(const ModuleName: string);
  var
    Item: TJclPeImportFuncItem;
    I, L: Integer;
    ImportLibItem: TJclPeImportLibItem;
    ExportItem: TJclPeExportFuncItem;
    ExportList: TJclPeExportFuncList;
  begin
    if Image.AttachedImage then
      LibPeDump.AttachLoadedModule(GetModuleHandle(PChar(ModuleName)))
    else
      LibPeDump.FileName := Image.ExpandModuleName(ModuleName);
    if not LibPeDump.StatusOK then
      Exit;
    ExportList := LibPeDump.ExportList;
    for L := 0 to Count - 1 do
    begin
      ImportLibItem := Items[L];
      if AnsiCompareText(ImportLibItem.Name, ModuleName) = 0 then
      begin
        for I := 0 to ImportLibItem.Count - 1 do
        begin
          Item := ImportLibItem[I];
          if Item.IsByOrdinal then
          begin
            ExportItem := ExportList.ItemFromOrdinal[Item.Ordinal];
            if (ExportItem <> nil) and (ExportItem.Name <> '') then
              Item.SetIndirectImportName(ExportItem.Name);
          end;
        end;
        ImportLibItem.SetSorted(False);
      end;
    end;
  end;

begin
  LibNamesList := TStringList.Create;
  try
    LibNamesList.Sorted := True;
    LibNamesList.Duplicates := dupIgnore;
    for L := 0 to Count - 1 do
      with Items[L] do
        for I := 0 to Count - 1 do
          if Items[I].IsByOrdinal then
            LibNamesList.Add(AnsiUpperCase(Name));
    LibPeDump := TJclPeImage.Create(True);
    try
      for I := 0 to LibNamesList.Count - 1 do
        TryGetNames(LibNamesList[I]);
    finally
      LibPeDump.Free;
    end;
    SortAllItemsList(FLastAllSortType, FLastAllSortDescending);
  finally
    LibNamesList.Free;
  end;
end;

//=== { TJclPeExportFuncItem } ===============================================

constructor TJclPeExportFuncItem.Create(AExportList: TJclPeExportFuncList;
  const AName, AForwardedName: string; AAddress: DWORD; AHint: Word;
  AOrdinal: Word; AResolveCheck: TJclPeResolveCheck);
var
  DotPos: Integer;
begin
  inherited Create;
  FExportList := AExportList;
  FName := AName;
  FForwardedName := AForwardedName;
  FAddress := AAddress;
  FHint := AHint;
  FOrdinal := AOrdinal;
  FResolveCheck := AResolveCheck;

  DotPos := AnsiPos('.', ForwardedName);
  if DotPos > 0 then
    FForwardedDotPos := Copy(ForwardedName, DotPos + 1, Length(ForwardedName) - DotPos)
  else
    FForwardedDotPos := '';
end;

function TJclPeExportFuncItem.GetAddressOrForwardStr: string;
begin
  if IsForwarded then
    Result := ForwardedName
  else
    FmtStr(Result, '%.8x', [Address]);
end;

function TJclPeExportFuncItem.GetForwardedFuncName: string;
begin
  if (Length(FForwardedDotPos) > 0) and (FForwardedDotPos[1] <> '#') then
    Result := FForwardedDotPos
  else
    Result := '';
end;

function TJclPeExportFuncItem.GetForwardedFuncOrdinal: DWORD;
begin
  if (Length(FForwardedDotPos) > 0) and (FForwardedDotPos[1] = '#') then
    Result := StrToIntDef(FForwardedDotPos, 0)
  else
    Result := 0;
end;

function TJclPeExportFuncItem.GetForwardedLibName: string;
begin
  if Length(FForwardedDotPos) = 0 then
    Result := ''
  else
    Result := AnsiLowerCase(Copy(FForwardedName, 1, Length(FForwardedName) - Length(FForwardedDotPos) - 1)) + BinaryExtensionLibrary;
end;

function TJclPeExportFuncItem.GetIsExportedVariable: Boolean;
begin
  case FExportList.Image.Target of
    taWin32:
    begin
      System.Error(rePlatformNotImplemented);//there is no BaseOfData in the 32-bit header for Win64
      Result := False;
      //Result := (Address >= FExportList.Image.OptionalHeader32.BaseOfData);
    end;
    taWin64:
      Result := False;
      // TODO equivalent for 64-bit modules
      //Result := (Address >= FExportList.Image.OptionalHeader64.BaseOfData);
  else
    Result := False;
  end;
end;

function TJclPeExportFuncItem.GetIsForwarded: Boolean;
begin
  Result := Length(FForwardedName) <> 0;
end;

function TJclPeExportFuncItem.GetMappedAddress: Pointer;
begin
  Result := FExportList.Image.RvaToVa(FAddress);
end;

function TJclPeExportFuncItem.GetSectionName: string;
begin
  if IsForwarded then
    Result := ''
  else
    with FExportList.Image do
      Result := ImageSectionNameFromRva[Address];
end;

procedure TJclPeExportFuncItem.SetResolveCheck(Value: TJclPeResolveCheck);
begin
  FResolveCheck := Value;
end;

// Export sort functions
function ExportSortByName(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeExportFuncItem(Item1).Name, TJclPeExportFuncItem(Item2).Name);
end;

function ExportSortByNameDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByName(Item2, Item1);
end;

function ExportSortByOrdinal(Item1, Item2: Pointer): Integer;
begin
  Result := TJclPeExportFuncItem(Item1).Ordinal - TJclPeExportFuncItem(Item2).Ordinal;
end;

function ExportSortByOrdinalDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByOrdinal(Item2, Item1);
end;

function ExportSortByHint(Item1, Item2: Pointer): Integer;
begin
  Result := TJclPeExportFuncItem(Item1).Hint - TJclPeExportFuncItem(Item2).Hint;
end;

function ExportSortByHintDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByHint(Item2, Item1);
end;

function ExportSortByAddress(Item1, Item2: Pointer): Integer;
begin
  Result := INT_PTR(TJclPeExportFuncItem(Item1).Address) - INT_PTR(TJclPeExportFuncItem(Item2).Address);
  if Result = 0 then
    Result := ExportSortByName(Item1, Item2);
end;

function ExportSortByAddressDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByAddress(Item2, Item1);
end;

function ExportSortByForwarded(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeExportFuncItem(Item1).ForwardedName, TJclPeExportFuncItem(Item2).ForwardedName);
  if Result = 0 then
    Result := ExportSortByName(Item1, Item2);
end;

function ExportSortByForwardedDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByForwarded(Item2, Item1);
end;

function ExportSortByAddrOrFwd(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeExportFuncItem(Item1).AddressOrForwardStr, TJclPeExportFuncItem(Item2).AddressOrForwardStr);
end;

function ExportSortByAddrOrFwdDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortByAddrOrFwd(Item2, Item1);
end;

function ExportSortBySection(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(TJclPeExportFuncItem(Item1).SectionName, TJclPeExportFuncItem(Item2).SectionName);
  if Result = 0 then
    Result := ExportSortByName(Item1, Item2);
end;

function ExportSortBySectionDESC(Item1, Item2: Pointer): Integer;
begin
  Result := ExportSortBySection(Item2, Item1);
end;

//=== { TJclPeExportFuncList } ===============================================

constructor TJclPeExportFuncList.Create(AImage: TJclPeImage);
begin
  inherited Create(AImage);
  FTotalResolveCheck := icNotChecked;
  CreateList;
end;

destructor TJclPeExportFuncList.Destroy;
begin
  FreeAndNil(FForwardedLibsList);
  inherited Destroy;
end;

function TJclPeExportFuncList.CanPerformFastNameSearch: Boolean;
begin
  Result := FSorted and (FLastSortType = esName) and not FLastSortDescending;
end;

procedure TJclPeExportFuncList.CheckForwards(PeImageCache: TJclPeImagesCache);
var
  I: Integer;
  FullFileName: TFileName;
  ForwardPeImage: TJclPeImage;
  ModuleResolveCheck: TJclPeResolveCheck;

  procedure PerformCheck(const ModuleName: string);
  var
    I: Integer;
    Item: TJclPeExportFuncItem;
    EL: TJclPeExportFuncList;
  begin
    EL := ForwardPeImage.ExportList;
    EL.PrepareForFastNameSearch;
    ModuleResolveCheck := icResolved;
    for I := 0 to Count - 1 do
    begin
      Item := Items[I];
      if (not Item.IsForwarded) or (Item.ResolveCheck <> icNotChecked) or
        (Item.ForwardedLibName <> ModuleName) then
        Continue;
      if EL.ItemFromName[Item.ForwardedFuncName] = nil then
      begin
        Item.SetResolveCheck(icUnresolved);
        ModuleResolveCheck := icUnresolved;
      end
      else
        Item.SetResolveCheck(icResolved);
    end;
  end;

begin
  if not AnyForwards then
    Exit;
  FTotalResolveCheck := icResolved;
  if PeImageCache <> nil then
    ForwardPeImage := nil // to make the compiler happy
  else
    ForwardPeImage := TJclPeImage.Create(True);
  try
    for I := 0 to ForwardedLibsList.Count - 1 do
    begin
      FullFileName := Image.ExpandModuleName(ForwardedLibsList[I]);
      if PeImageCache <> nil then
        ForwardPeImage := PeImageCache[FullFileName]
      else
        ForwardPeImage.FileName := FullFileName;
      if ForwardPeImage.StatusOK then
        PerformCheck(ForwardedLibsList[I])
      else
        ModuleResolveCheck := icUnresolved;
      FForwardedLibsList.Objects[I] := Pointer(ModuleResolveCheck);
      if ModuleResolveCheck = icUnresolved then
        FTotalResolveCheck := icUnresolved;
    end;
  finally
    if PeImageCache = nil then
      ForwardPeImage.Free;
  end;
end;

procedure TJclPeExportFuncList.CreateList;
var
  Functions: Pointer;
  Address, NameCount: DWORD;
  NameOrdinals: PWORD;
  Names: PDWORD;
  I: Integer;
  ExportItem: TJclPeExportFuncItem;
  ExportVABegin, ExportVAEnd: DWORD;
  UTF8Name: AnsiString;
  ForwardedName, ExportName: string;
begin
  with Image do
  begin
    if not StatusOK then
      Exit;
    with Directories[IMAGE_DIRECTORY_ENTRY_EXPORT] do
    begin
      ExportVABegin := VirtualAddress;
      ExportVAEnd := VirtualAddress + NativeInt(Size);
    end;
    FExportDir := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_EXPORT);
    if FExportDir <> nil then
    begin
      FBase := FExportDir^.Base;
      FFunctionCount := FExportDir^.NumberOfFunctions;
      Functions := RvaToVa(FExportDir^.AddressOfFunctions);
      NameOrdinals := RvaToVa(FExportDir^.AddressOfNameOrdinals);
      Names := RvaToVa(FExportDir^.AddressOfNames);
      NameCount := FExportDir^.NumberOfNames;
      Count := FExportDir^.NumberOfFunctions;

      for I := 0 to Count - 1 do
      begin
        Address := PDWORD(NativeInt(Functions) + NativeInt(I) * SizeOf(DWORD))^;
        if (Address >= ExportVABegin) and (Address <= ExportVAEnd) then
        begin
          FAnyForwards := True;
          UTF8Name := PAnsiChar(RvaToVa(Address));
          ForwardedName := string(UTF8Name);
        end
        else
          ForwardedName := '';

        ExportItem := TJclPeExportFuncItem.Create(Self, '',
          ForwardedName, Address, $FFFF, NativeInt(I) + FBase, icNotChecked);

        List[I] := ExportItem;
      end;

      for I := 0 to NameCount - 1 do
      begin
          // named function
        UTF8Name := PAnsiChar(RvaToVa(Names^));
        ExportName := string(UTF8Name);

        ExportItem := TJclPeExportFuncItem(List[NameOrdinals^]);
        ExportItem.FName := ExportName;
        ExportItem.FHint := I;

        Inc(NameOrdinals);
        Inc(Names);
      end;
    end;
  end;
end;

function TJclPeExportFuncList.GetForwardedLibsList: TStrings;
var
  I: Integer;
begin
  if FForwardedLibsList = nil then
  begin
    FForwardedLibsList := TStringList.Create;
    FForwardedLibsList.Sorted := True;
    FForwardedLibsList.Duplicates := dupIgnore;
    if FAnyForwards then
      for I := 0 to Count - 1 do
        with Items[I] do
          if IsForwarded then
            FForwardedLibsList.AddObject(ForwardedLibName, Pointer(icNotChecked));
  end;
  Result := FForwardedLibsList;
end;

function TJclPeExportFuncList.GetItemFromAddress(Address: DWORD): TJclPeExportFuncItem;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to Count - 1 do
    if Items[I].Address = Address then
    begin
      Result := Items[I];
      Break;
    end;
end;

function TJclPeExportFuncList.GetItemFromName(const Name: string): TJclPeExportFuncItem;
var
  L, H, I, C: Integer;
  B: Boolean;
begin
  Result := nil;
  if CanPerformFastNameSearch then
  begin
    L := 0;
    H := Count - 1;
    B := False;
    while L <= H do
    begin
      I := (L + H) shr 1;
      C := CompareStr(Items[I].Name, Name);
      if C < 0 then
        L := I + 1
      else
      begin
        H := I - 1;
        if C = 0 then
        begin
          B := True;
          L := I;
        end;
      end;
    end;
    if B then
      Result := Items[L];
  end
  else
    for I := 0 to Count - 1 do
      if Items[I].Name = Name then
      begin
        Result := Items[I];
        Break;
      end;
end;

function TJclPeExportFuncList.GetItemFromOrdinal(Ordinal: DWORD): TJclPeExportFuncItem;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to Count - 1 do
    if Items[I].Ordinal = Ordinal then
    begin
      Result := Items[I];
      Break;
    end;
end;

function TJclPeExportFuncList.GetItems(Index: Integer): TJclPeExportFuncItem;
begin
  Result := TJclPeExportFuncItem(Get(Index));
end;

function TJclPeExportFuncList.GetName: string;
var
  UTF8ExportName: AnsiString;
begin
  if (FExportDir = nil) or (FExportDir^.Name = 0) then
    Result := ''
  else
  begin
    UTF8ExportName := PAnsiChar(Image.RvaToVa(FExportDir^.Name));
    Result := string(UTF8ExportName);
  end;
end;

class function TJclPeExportFuncList.ItemName(Item: TJclPeExportFuncItem): string;
begin
  if Item = nil then
    Result := ''
  else
    Result := Item.Name;
end;

function TJclPeExportFuncList.OrdinalValid(Ordinal: DWORD): Boolean;
begin
  Result := (FExportDir <> nil) and (Ordinal >= Base) and
    (Ordinal < FunctionCount + Base);
end;

procedure TJclPeExportFuncList.PrepareForFastNameSearch;
begin
  if not CanPerformFastNameSearch then
    SortList(esName, False);
end;

function TJclPeExportFuncList.SmartFindName(const CompareName: string;
  Options: TJclSmartCompOptions): TJclPeExportFuncItem;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to Count - 1 do
  begin
    if PeSmartFunctionNameSame(CompareName, Items[I].Name, Options) then
    begin
      Result := Items[I];
      Break;
    end;
  end;
end;

procedure TJclPeExportFuncList.SortList(SortType: TJclPeExportSort; Descending: Boolean);
const
  SortFunctions: array [TJclPeExportSort, Boolean] of TListSortCompare =
    ((ExportSortByName, ExportSortByNameDESC),
     (ExportSortByOrdinal, ExportSortByOrdinalDESC),
     (ExportSortByHint, ExportSortByHintDESC),
     (ExportSortByAddress, ExportSortByAddressDESC),
     (ExportSortByForwarded, ExportSortByForwardedDESC),
     (ExportSortByAddrOrFwd, ExportSortByAddrOrFwdDESC),
     (ExportSortBySection, ExportSortBySectionDESC)
    );
begin
  if not FSorted or (SortType <> FLastSortType) or (Descending <> FLastSortDescending) then
  begin
    Sort(SortFunctions[SortType, Descending]);
    FLastSortType := SortType;
    FLastSortDescending := Descending;
    FSorted := True;
  end;
end;

//=== { TJclPeResourceRawStream } ============================================

constructor TJclPeResourceRawStream.Create(AResourceItem: TJclPeResourceItem);
begin
  Assert(not AResourceItem.IsDirectory);
  inherited Create;
  SetPointer(AResourceItem.RawEntryData, AResourceItem.RawEntryDataSize);
end;

function TJclPeResourceRawStream.Write(const Buffer; Count: Integer): Longint;
begin
  raise EJclPeImageError.Create('Stream is read-only');
end;

//=== { TJclPeResourceItem } =================================================

constructor TJclPeResourceItem.Create(AImage: TJclPeImage;
  AParentItem: TJclPeResourceItem; AEntry: PImageResourceDirectoryEntry);
begin
  inherited Create;
  FImage := AImage;
  FEntry := AEntry;
  FParentItem := AParentItem;
  if AParentItem = nil then
    FLevel := 1
  else
    FLevel := AParentItem.Level + 1;
end;

destructor TJclPeResourceItem.Destroy;
begin
  FreeAndNil(FList);
  inherited Destroy;
end;

function TJclPeResourceItem.CompareName(AName: PChar): Boolean;
var
  P: PChar;
begin
  if IsName then
    P := PChar(Name)
  else
    P := PChar(FEntry^.Name and $FFFF); // Integer encoded in a PChar
  Result := CompareResourceName(AName, P);
end;

function TJclPeResourceItem.GetDataEntry: PImageResourceDataEntry;
begin
  if GetIsDirectory then
    Result := nil
  else
    Result := PImageResourceDataEntry(OffsetToRawData(FEntry^.OffsetToData));
end;

function TJclPeResourceItem.GetIsDirectory: Boolean;
begin
  Result := FEntry^.OffsetToData and IMAGE_RESOURCE_DATA_IS_DIRECTORY <> 0;
end;

function TJclPeResourceItem.GetIsName: Boolean;
begin
  Result := FEntry^.Name and IMAGE_RESOURCE_NAME_IS_STRING <> 0;
end;

function TJclPeResourceItem.GetLangID: LANGID;
begin
  if IsDirectory then
  begin
    GetList;
    if FList.Count = 1 then
      Result := StrToIntDef(FList[0].Name, 0)
    else
      Result := 0;
  end
  else
    Result := StrToIntDef(Name, 0);
end;

function TJclPeResourceItem.GetList: TJclPeResourceList;
begin
  if not IsDirectory then
  begin
    if Image.NoExceptions then
    begin
      Result := nil;
      Exit;
    end
    else
      raise EJclPeImageError.Create('Not a resource directory');
  end;
  if FList = nil then
    FList := FImage.ResourceListCreate(SubDirData, Self);
  Result := FList;
end;

function TJclPeResourceItem.GetName: string;
begin
  if IsName then
  begin
    if FNameCache = '' then
    begin
      with PImageResourceDirStringU(OffsetToRawData(FEntry^.Name))^ do
        FNameCache := WideCharLenToString(NameString, Length);
      FNameCache := FNameCache;
    end;
    Result := FNameCache;
  end
  else
    Result := IntToStr(FEntry^.Name and $FFFF);
end;

function TJclPeResourceItem.GetParameterName: string;
begin
  if IsName then
    Result := Name
  else
    Result := Format('#%d', [FEntry^.Name and $FFFF]);
end;

function TJclPeResourceItem.GetRawEntryData: Pointer;
begin
  if GetIsDirectory then
    Result := nil
  else
    Result := FImage.RvaToVa(GetDataEntry^.OffsetToData);
end;

function TJclPeResourceItem.GetRawEntryDataSize: Integer;
begin
  if GetIsDirectory then
    Result := -1
  else
    Result := PImageResourceDataEntry(OffsetToRawData(FEntry^.OffsetToData))^.Size;
end;

function TJclPeResourceItem.GetResourceType: TJclPeResourceKind;
begin
  with Level1Item do
  begin
    if FEntry^.Name < Cardinal(High(TJclPeResourceKind)) then
      Result := TJclPeResourceKind(FEntry^.Name)
    else
      Result := rtUserDefined
  end;
end;

function TJclPeResourceItem.GetResourceTypeStr: string;
begin
  with Level1Item do
  begin
    if FEntry^.Name < Cardinal(High(TJclPeResourceKind)) then
      Result := Copy(GetEnumName(TypeInfo(TJclPeResourceKind), Ord(FEntry^.Name)), 3, 30)
    else
      Result := Name;
  end;
end;

function TJclPeResourceItem.Level1Item: TJclPeResourceItem;
begin
  Result := Self;
  while Result.FParentItem <> nil do
    Result := Result.FParentItem;
end;

function TJclPeResourceItem.OffsetToRawData(Ofs: DWORD): NativeInt;
begin
  Result := (Ofs and $7FFFFFFF) + Image.ResourceVA;
end;

function TJclPeResourceItem.SubDirData: PImageResourceDirectory;
begin
  Result := Pointer(OffsetToRawData(FEntry^.OffsetToData));
end;

//=== { TJclPeResourceList } =================================================

constructor TJclPeResourceList.Create(AImage: TJclPeImage;
  AParentItem: TJclPeResourceItem; ADirectory: PImageResourceDirectory);
begin
  inherited Create(AImage);
  FDirectory := ADirectory;
  FParentItem := AParentItem;
  CreateList(AParentItem);
end;

procedure TJclPeResourceList.CreateList(AParentItem: TJclPeResourceItem);
var
  Entry: PImageResourceDirectoryEntry;
  DirItem: TJclPeResourceItem;
  I: Integer;
begin
  if FDirectory = nil then
    Exit;
  Entry := Pointer(NativeInt(FDirectory) + SizeOf(TImageResourceDirectory));
  for I := 1 to DWORD(FDirectory^.NumberOfNamedEntries) + DWORD(FDirectory^.NumberOfIdEntries) do
  begin
    DirItem := Image.ResourceItemCreate(Entry, AParentItem);
    Add(DirItem);
    Inc(Entry);
  end;
end;

function TJclPeResourceList.FindName(const Name: string): TJclPeResourceItem;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to Count - 1 do
    if Trim(Items[I].Name) = Trim(Name) then
    begin
      Result := Items[I];
      Break;
    end;
end;

function TJclPeResourceList.GetItems(Index: Integer): TJclPeResourceItem;
begin
  Result := TJclPeResourceItem(Get(Index));
end;

//=== { TJclPeRootResourceList } =============================================

destructor TJclPeRootResourceList.Destroy;
begin
  FreeAndNil(FManifestContent);
  inherited Destroy;
end;

function TJclPeRootResourceList.FindResource(ResourceType: TJclPeResourceKind;
  const ResourceName: string): TJclPeResourceItem;
var
  I: Integer;
  TypeItem: TJclPeResourceItem;
begin
  Result := nil;
  TypeItem := nil;
  for I := 0 to Count - 1 do
  begin
    if Items[I].ResourceType = ResourceType then
    begin
      TypeItem := Items[I];
      Break;
    end;
  end;
  if TypeItem <> nil then
    if ResourceName = '' then
      Result := TypeItem
    else
      with TypeItem.List do
        for I := 0 to Count - 1 do
          if Items[I].Name = ResourceName then
          begin
            Result := Items[I];
            Break;
          end;
end;

function TJclPeRootResourceList.FindResource(const ResourceType: PChar;
  const ResourceName: PChar): TJclPeResourceItem;
var
  I: Integer;
  TypeItem: TJclPeResourceItem;
begin
  Result := nil;
  TypeItem := nil;
  for I := 0 to Count - 1 do
    if Items[I].CompareName(ResourceType) then
    begin
      TypeItem := Items[I];
      Break;
    end;
  if TypeItem <> nil then
    if ResourceName = nil then
      Result := TypeItem
    else
      with TypeItem.List do
        for I := 0 to Count - 1 do
          if Items[I].CompareName(ResourceName) then
          begin
            Result := Items[I];
            Break;
          end;
end;

function TJclPeRootResourceList.GetManifestContent: TStrings;
var
  ManifestFileName: string;
  ResItem: TJclPeResourceItem;
  ResStream: TJclPeResourceRawStream;
begin
  if FManifestContent = nil then
  begin
    FManifestContent := TStringList.Create;
    ResItem := FindResource(RT_MANIFEST, CREATEPROCESS_MANIFEST_RESOURCE_ID);
    if ResItem = nil then
    begin
      ManifestFileName := Image.FileName + MANIFESTExtension;
      if FileExists(ManifestFileName) then
        FManifestContent.LoadFromFile(ManifestFileName);
    end
    else
    begin
      ResStream := TJclPeResourceRawStream.Create(ResItem.List[0]);
      try
        FManifestContent.LoadFromStream(ResStream);
      finally
        ResStream.Free;
      end;
    end;
  end;
  Result := FManifestContent;
end;

function TJclPeRootResourceList.ListResourceNames(ResourceType: TJclPeResourceKind;
  const Strings: TStrings): Boolean;
var
  ResTypeItem, TempItem: TJclPeResourceItem;
  I: Integer;
begin
  ResTypeItem := FindResource(ResourceType, '');
  Result := (ResTypeItem <> nil);
  if Result then
  begin
    Strings.BeginUpdate;
    try
      with ResTypeItem.List do
        for I := 0 to Count - 1 do
        begin
          TempItem := Items[I];
          Strings.AddObject(TempItem.Name, Pointer(TempItem.IsName));
        end;
    finally
      Strings.EndUpdate;
    end;
  end;
end;

//=== { TJclPeRelocEntry } ===================================================

constructor TJclPeRelocEntry.Create(AChunk: PImageBaseRelocation; ACount: Integer);
begin
  inherited Create;
  FChunk := AChunk;
  FCount := ACount;
end;

function TJclPeRelocEntry.GetRelocations(Index: Integer): TJclPeRelocation;
var
  Temp: Word;
begin
  Temp := PWord(NativeInt(FChunk) + SizeOf(TImageBaseRelocation) + DWORD(Index) * SizeOf(Word))^;
  Result.Address := Temp and $0FFF;
  Result.RelocType := (Temp and $F000) shr 12;
  Result.VirtualAddress := NativeInt(Result.Address) + VirtualAddress;
end;

function TJclPeRelocEntry.GetSize: DWORD;
begin
  Result := FChunk^.SizeOfBlock;
end;

function TJclPeRelocEntry.GetVirtualAddress: DWORD;
begin
  Result := FChunk^.VirtualAddress;
end;

//=== { TJclPeRelocList } ====================================================

constructor TJclPeRelocList.Create(AImage: TJclPeImage);
begin
  inherited Create(AImage);
  CreateList;
end;

procedure TJclPeRelocList.CreateList;
var
  Chunk: PImageBaseRelocation;
  Item: TJclPeRelocEntry;
  RelocCount: Integer;
begin
  with Image do
  begin
    if not StatusOK then
      Exit;
    Chunk := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_BASERELOC);
    if Chunk = nil then
      Exit;
    FAllItemCount := 0;
    while Chunk^.SizeOfBlock <> 0 do
    begin
      RelocCount := (Chunk^.SizeOfBlock - SizeOf(TImageBaseRelocation)) div SizeOf(Word);
      Item := TJclPeRelocEntry.Create(Chunk, RelocCount);
      Inc(FAllItemCount, RelocCount);
      Add(Item);
      Chunk := Pointer(NativeInt(Chunk) + Chunk^.SizeOfBlock);
    end;
  end;
end;

function TJclPeRelocList.GetAllItems(Index: Integer): TJclPeRelocation;
var
  I, N, C: Integer;
begin
  N := Index;
  for I := 0 to Count - 1 do
  begin
    C := Items[I].Count;
    Dec(N, C);
    if N < 0 then
    begin
      Result := Items[I][N + C];
      Break;
    end;
  end;
end;

function TJclPeRelocList.GetItems(Index: Integer): TJclPeRelocEntry;
begin
  Result := TJclPeRelocEntry(Get(Index));
end;

//=== { TJclPeDebugList } ====================================================

constructor TJclPeDebugList.Create(AImage: TJclPeImage);
begin
  inherited Create(AImage);
  OwnsObjects := False;
  CreateList;
end;

procedure TJclPeDebugList.CreateList;
var
  DebugImageDir: TImageDataDirectory;
  DebugDir: PImageDebugDirectory;
  Header: PImageSectionHeader;
  FormatCount, I: Integer;
begin
  with Image do
  begin
    if not StatusOK then
      Exit;
    DebugImageDir := Directories[IMAGE_DIRECTORY_ENTRY_DEBUG];
    if DebugImageDir.VirtualAddress = 0 then
      Exit;
    if GetSectionHeader(DebugSectionName, Header) and
      (Header^.VirtualAddress = DebugImageDir.VirtualAddress) then
    begin
      FormatCount := DebugImageDir.Size;
      DebugDir := RvaToVa(Header^.VirtualAddress);
    end
    else
    begin
      if not GetSectionHeader(ReadOnlySectionName, Header) then
        Exit;
      FormatCount := DebugImageDir.Size div SizeOf(TImageDebugDirectory);
      DebugDir := Pointer(MappedAddress + DebugImageDir.VirtualAddress -
        Header^.VirtualAddress + Header^.PointerToRawData);
    end;
    for I := 1 to FormatCount do
    begin
      Add(TObject(DebugDir));
      Inc(DebugDir);
    end;
  end;
end;

function TJclPeDebugList.GetItems(Index: Integer): TImageDebugDirectory;
begin
  Result := PImageDebugDirectory(Get(Index))^;
end;

//=== { TJclPeCertificate } ==================================================

constructor TJclPeCertificate.Create(AHeader: TWinCertificate; AData: Pointer);
begin
  inherited Create;
  FHeader := AHeader;
  FData := AData;
end;

//=== { TJclPeCertificateList } ==============================================

constructor TJclPeCertificateList.Create(AImage: TJclPeImage);
begin
  inherited Create(AImage);
  CreateList;
end;

procedure TJclPeCertificateList.CreateList;
var
  Directory: TImageDataDirectory;
  CertPtr: PChar;
  TotalSize: Integer;
  Item: TJclPeCertificate;
begin
  Directory := Image.Directories[IMAGE_DIRECTORY_ENTRY_SECURITY];
  if Directory.VirtualAddress = 0 then
    Exit;
  CertPtr := Image.RawToVa(Directory.VirtualAddress); // Security directory is a raw offset
  TotalSize := Directory.Size;
  while TotalSize >= SizeOf(TWinCertificate) do
  begin
    Item := TJclPeCertificate.Create(PWinCertificate(CertPtr)^, CertPtr + SizeOf(TWinCertificate));
    Dec(TotalSize, Item.Header.dwLength);
    Add(Item);
  end;
end;

function TJclPeCertificateList.GetItems(Index: Integer): TJclPeCertificate;
begin
  Result := TJclPeCertificate(Get(Index));
end;

//=== { TJclPeCLRHeader } ====================================================

constructor TJclPeCLRHeader.Create(AImage: TJclPeImage);
begin
  FImage := AImage;
  ReadHeader;
end;

function TJclPeCLRHeader.GetHasMetadata: Boolean;
const
  METADATA_SIGNATURE = $424A5342; // Reference: Partition II Metadata.doc - 23.2.1 Metadata root
begin
  with Header.MetaData do
    Result := (VirtualAddress <> 0) and (PDWORD(FImage.RvaToVa(VirtualAddress))^ = METADATA_SIGNATURE);
end;
{ TODO -cDOC : "Flier Lu" <flier_lu att yahoo dott com dott cn> }

function TJclPeCLRHeader.GetVersionString: string;
begin
  //Result := FormatVersionString(Header.MajorRuntimeVersion, Header.MinorRuntimeVersion);
end;

procedure TJclPeCLRHeader.ReadHeader;
var
  HeaderPtr: PImageCor20Header;
begin
  HeaderPtr := Image.DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR);
  if (HeaderPtr <> nil) and (HeaderPtr^.cb >= SizeOf(TImageCor20Header)) then
    FHeader := HeaderPtr^;
end;

//=== { TJclPeImage } ========================================================

constructor TJclPeImage.Create(ANoExceptions: Boolean);
begin
  FNoExceptions := ANoExceptions;
  FReadOnlyAccess := True;
  FImageSections := TStringList.Create;
  FStringTable := TStringList.Create;
end;

destructor TJclPeImage.Destroy;
begin
  Clear;
  FreeAndNil(FImageSections);
  FStringTable.Free;

  inherited Destroy;
end;

procedure TJclPeImage.AfterOpen;
begin
end;

procedure TJclPeImage.AttachLoadedModule(const Handle: HMODULE);
  procedure AttachLoadedModule32;
  var
    NtHeaders: PImageNtHeaders32;
  begin
    NtHeaders := PeMapImgNtHeaders32(Pointer(Handle));
    if NtHeaders = nil then
      FStatus := stNotPE
    else
    begin
      FStatus := stOk;
      FAttachedImage := True;
      FFileName := '';//GetModulePath(Handle);
      // OF: possible loss of data
      FLoadedImage.ModuleName := PAnsiChar(AnsiString(FFileName));
      FLoadedImage.hFile := INVALID_HANDLE_VALUE;
      FLoadedImage.MappedAddress := Pointer(Handle);
      FLoadedImage.FileHeader := PImageNtHeaders(NtHeaders);
      FLoadedImage.NumberOfSections := NtHeaders^.FileHeader.NumberOfSections;
      FLoadedImage.Sections := PeMapImgSections32(NtHeaders);
      FLoadedImage.LastRvaSection := FLoadedImage.Sections;
      FLoadedImage.Characteristics := NtHeaders^.FileHeader.Characteristics;
      FLoadedImage.fSystemImage := (FLoadedImage.Characteristics and IMAGE_FILE_SYSTEM <> 0);
      FLoadedImage.fDOSImage := False;
      FLoadedImage.SizeOfImage := NtHeaders^.OptionalHeader.SizeOfImage;
      ReadImageSections;
      ReadStringTable;
      AfterOpen;
    end;
    RaiseStatusException;
  end;

  procedure AttachLoadedModule64;
   var
    NtHeaders: PImageNtHeaders64;
  begin
    NtHeaders := PeMapImgNtHeaders64(Pointer(Handle));
    if NtHeaders = nil then
      FStatus := stNotPE
    else
    begin
      FStatus := stOk;
      FAttachedImage := True;
      FFileName := '';//GetModulePath(Handle);
      // OF: possible loss of data
      FLoadedImage.ModuleName := PAnsiChar(AnsiString(FFileName));
      FLoadedImage.hFile := INVALID_HANDLE_VALUE;
      FLoadedImage.MappedAddress := Pointer(Handle);
      FLoadedImage.FileHeader := PImageNtHeaders(NtHeaders);
      FLoadedImage.NumberOfSections := NtHeaders^.FileHeader.NumberOfSections;
      FLoadedImage.Sections := PeMapImgSections64(NtHeaders);
      FLoadedImage.LastRvaSection := FLoadedImage.Sections;
      FLoadedImage.Characteristics := NtHeaders^.FileHeader.Characteristics;
      FLoadedImage.fSystemImage := (FLoadedImage.Characteristics and IMAGE_FILE_SYSTEM <> 0);
      FLoadedImage.fDOSImage := False;
      FLoadedImage.SizeOfImage := NtHeaders^.OptionalHeader.SizeOfImage;
      ReadImageSections;
      ReadStringTable;
      AfterOpen;
    end;
    RaiseStatusException;
  end;
begin
  Clear;
  if Handle = 0 then
    Exit;
  FTarget := PeMapImgTarget(Pointer(Handle));
  case Target of
    taWin32:
      AttachLoadedModule32;
    taWin64:
      AttachLoadedModule64;
    taUnknown:
      FStatus := stNotSupported;
  end;
end;

function TJclPeImage.CalculateCheckSum: DWORD;
var
  C: DWORD;
begin
  if StatusOK then
  begin
    CheckNotAttached;
    if CheckSumMappedFile(FLoadedImage.MappedAddress, FLoadedImage.SizeOfImage,
      C, Result) = nil then
        RaiseLastOSError;
  end
  else
    Result := 0;
end;

procedure TJclPeImage.CheckNotAttached;
begin
  if FAttachedImage then
    raise EJclPeImageError.Create('Feature is not available for attached images');
end;

procedure TJclPeImage.Clear;
begin
  FImageSections.Clear;
  FStringTable.Clear;
  FreeAndNil(FCertificateList);
  FreeAndNil(FCLRHeader);
  FreeAndNil(FDebugList);
  FreeAndNil(FImportList);
  FreeAndNil(FExportList);
  FreeAndNil(FRelocationList);
  FreeAndNil(FResourceList);
  if not FAttachedImage and StatusOK then
    UnMapAndLoad(FLoadedImage);
  FillChar(FLoadedImage, SizeOf(FLoadedImage), #0);
  FStatus := stNotLoaded;
  FAttachedImage := False;
end;

class function TJclPeImage.DateTimeToStamp(const DateTime: TDateTime): DWORD;
begin
  Result := 0;
  //Result := Round((DateTime - UnixTimeStart) * SecsPerDay);
end;

class function TJclPeImage.DebugTypeNames(DebugType: DWORD): string;
begin
  case DebugType of
    IMAGE_DEBUG_TYPE_UNKNOWN:
      Result := 'UNKNOWN';
    IMAGE_DEBUG_TYPE_COFF:
      Result := 'COFF';
    IMAGE_DEBUG_TYPE_CODEVIEW:
      Result := 'CODEVIEW';
    IMAGE_DEBUG_TYPE_FPO:
      Result := 'FPO';
    IMAGE_DEBUG_TYPE_MISC:
      Result := 'MISC';
    IMAGE_DEBUG_TYPE_EXCEPTION:
      Result := 'EXCEPTION';
    IMAGE_DEBUG_TYPE_FIXUP:
      Result := 'FIXUP';
    IMAGE_DEBUG_TYPE_OMAP_TO_SRC:
      Result := 'OMAP_TO_SRC';
    IMAGE_DEBUG_TYPE_OMAP_FROM_SRC:
      Result := 'OMAP_FROM_SRC';
  else
    Result := 'UNKNOWN';
  end;
end;

function TJclPeImage.DirectoryEntryToData(Directory: Word): Pointer;
var
  Size: DWORD;
begin
  Size := 0;
  Result := ImageDirectoryEntryToData(FLoadedImage.MappedAddress, FAttachedImage, Directory, Size);
end;

class function TJclPeImage.ExpandBySearchPath(const ModuleName, BasePath: string): TFileName;
var
  FullName: array [0..MAX_PATH] of Char;
  FilePart: PChar;
begin
  Result := IncludeTrailingPathDelimiter(ExtractFilePath(BasePath)) + ModuleName;
  if FileExists(Result) then
    Exit;
  FilePart := nil;
  if SearchPath(nil, PChar(ModuleName), nil, Length(FullName), FullName, FilePart) = 0 then
    Result := ModuleName
  else
    Result := FullName;
end;

function TJclPeImage.ExpandModuleName(const ModuleName: string): TFileName;
begin
  Result := ExpandBySearchPath(ModuleName, ExtractFilePath(FFileName));
end;

function TJclPeImage.GetCertificateList: TJclPeCertificateList;
begin
  if FCertificateList = nil then
    FCertificateList := TJclPeCertificateList.Create(Self);
  Result := FCertificateList;
end;

function TJclPeImage.GetCLRHeader: TJclPeCLRHeader;
begin
  if FCLRHeader = nil then
    FCLRHeader := TJclPeCLRHeader.Create(Self);
  Result := FCLRHeader;
end;

function TJclPeImage.GetDebugList: TJclPeDebugList;
begin
  if FDebugList = nil then
    FDebugList := TJclPeDebugList.Create(Self);
  Result := FDebugList;
end;

function TJclPeImage.GetDescription: string;
var
  UTF8DescriptionName: AnsiString;
begin
  if DirectoryExists[IMAGE_DIRECTORY_ENTRY_COPYRIGHT] then
  begin
    UTF8DescriptionName := PAnsiChar(DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_COPYRIGHT));
    Result := string(UTF8DescriptionName);
  end
  else
    Result := '';
end;

function TJclPeImage.GetDirectories(Directory: Word): TImageDataDirectory;
begin
  if StatusOK then
  begin
    case Target of
      taWin32:
        Result := PImageNtHeaders32(FLoadedImage.FileHeader)^.OptionalHeader.DataDirectory[Directory];
      taWin64:
        Result := PImageNtHeaders64(FLoadedImage.FileHeader)^.OptionalHeader.DataDirectory[Directory];
    else
      Result.VirtualAddress := 0;
      Result.Size := 0;
    end
  end
  else
  begin
    Result.VirtualAddress := 0;
    Result.Size := 0;
  end;
end;

function TJclPeImage.GetDirectoryExists(Directory: Word): Boolean;
begin
  Result := (Directories[Directory].VirtualAddress <> 0);
end;

function TJclPeImage.GetExportList: TJclPeExportFuncList;
begin
  if FExportList = nil then
    FExportList := TJclPeExportFuncList.Create(Self);
  Result := FExportList;
end;

function TJclPeImage.GetImageSectionCount: Integer;
begin
  Result := FImageSections.Count;
end;

function TJclPeImage.GetImageSectionFullNames(Index: Integer): string;
var
  Offset: Integer;
begin
  Result := ImageSectionNames[Index];
  if (Length(Result) > 0) and (Result[1] = '/') and TryStrToInt(Copy(Result, 2, MaxInt), Offset) then
    Result := GetNameInStringTable(Offset);
end;

function TJclPeImage.GetImageSectionHeaders(Index: Integer): TImageSectionHeader;
begin
  Result := PImageSectionHeader(FImageSections.Objects[Index])^;
end;

function TJclPeImage.GetImageSectionNameFromRva(const Rva: DWORD): string;
begin
  Result := GetSectionName(RvaToSection(Rva));
end;

function TJclPeImage.GetImageSectionNames(Index: Integer): string;
begin
  Result := FImageSections[Index];
end;

function TJclPeImage.GetImportList: TJclPeImportList;
begin
  if FImportList = nil then
    FImportList := TJclPeImportList.Create(Self);
  Result := FImportList;
end;

function TJclPeImage.GetLoadConfigValues(Index: TJclLoadConfig): string;
  function GetLoadConfigValues32(Index: TJclLoadConfig): string;
  var
    LoadConfig: PIMAGE_LOAD_CONFIG_DIRECTORY32;
  begin
    LoadConfig := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG);
    if LoadConfig <> nil then
      with LoadConfig^ do
        case Index of
          JclLoadConfig_Characteristics:
            Result := IntToHex(Size, 8);
          JclLoadConfig_TimeDateStamp:
            Result := IntToHex(TimeDateStamp, 8);
          //JclLoadConfig_Version:
          //  Result := FormatVersionString(MajorVersion, MinorVersion);
          JclLoadConfig_GlobalFlagsClear:
            Result := IntToHex(GlobalFlagsClear, 8);
          JclLoadConfig_GlobalFlagsSet:
            Result := IntToHex(GlobalFlagsSet, 8);
          JclLoadConfig_CriticalSectionDefaultTimeout:
            Result := IntToHex(CriticalSectionDefaultTimeout, 8);
          JclLoadConfig_DeCommitFreeBlockThreshold:
            Result := IntToHex(DeCommitFreeBlockThreshold, 8);
          JclLoadConfig_DeCommitTotalFreeThreshold:
            Result := IntToHex(DeCommitTotalFreeThreshold, 8);
          JclLoadConfig_LockPrefixTable:
            Result := IntToHex(LockPrefixTable, 8);
          JclLoadConfig_MaximumAllocationSize:
            Result := IntToHex(MaximumAllocationSize, 8);
          JclLoadConfig_VirtualMemoryThreshold:
            Result := IntToHex(VirtualMemoryThreshold, 8);
          JclLoadConfig_ProcessHeapFlags:
            Result := IntToHex(ProcessHeapFlags, 8);
          JclLoadConfig_ProcessAffinityMask:
            Result := IntToHex(ProcessAffinityMask, 8);
          JclLoadConfig_CSDVersion:
            Result := IntToHex(CSDVersion, 4);
          JclLoadConfig_Reserved1:
            Result := IntToHex(Reserved1, 4);
          JclLoadConfig_EditList:
            Result := IntToHex(EditList, 8);
          JclLoadConfig_Reserved:
            Result := 'Reserved';
        end;
  end;
  function GetLoadConfigValues64(Index: TJclLoadConfig): string;
  var
    LoadConfig: PIMAGE_LOAD_CONFIG_DIRECTORY64;
  begin
    LoadConfig := DirectoryEntryToData(IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG);
    if LoadConfig <> nil then
      with LoadConfig^ do
        case Index of
          JclLoadConfig_Characteristics:
            Result := IntToHex(Size, 8);
          JclLoadConfig_TimeDateStamp:
            Result := IntToHex(TimeDateStamp, 8);
          //JclLoadConfig_Version:
          //  Result := FormatVersionString(MajorVersion, MinorVersion);
          JclLoadConfig_GlobalFlagsClear:
            Result := IntToHex(GlobalFlagsClear, 8);
          JclLoadConfig_GlobalFlagsSet:
            Result := IntToHex(GlobalFlagsSet, 8);
          JclLoadConfig_CriticalSectionDefaultTimeout:
            Result := IntToHex(CriticalSectionDefaultTimeout, 8);
          JclLoadConfig_DeCommitFreeBlockThreshold:
            Result := IntToHex(DeCommitFreeBlockThreshold, 16);
          JclLoadConfig_DeCommitTotalFreeThreshold:
            Result := IntToHex(DeCommitTotalFreeThreshold, 16);
          JclLoadConfig_LockPrefixTable:
            Result := IntToHex(LockPrefixTable, 16);
          JclLoadConfig_MaximumAllocationSize:
            Result := IntToHex(MaximumAllocationSize, 16);
          JclLoadConfig_VirtualMemoryThreshold:
            Result := IntToHex(VirtualMemoryThreshold, 16);
          JclLoadConfig_ProcessHeapFlags:
            Result := IntToHex(ProcessHeapFlags, 8);
          JclLoadConfig_ProcessAffinityMask:
            Result := IntToHex(ProcessAffinityMask, 16);
          JclLoadConfig_CSDVersion:
            Result := IntToHex(CSDVersion, 4);
          JclLoadConfig_Reserved1:
            Result := IntToHex(Reserved1, 4);
          JclLoadConfig_EditList:
            Result := IntToHex(EditList, 16);
          JclLoadConfig_Reserved:
            Result := 'Reserved';
        end;
  end;
begin
  Result := '';
  case Target of
    taWin32:
      Result := GetLoadConfigValues32(Index);
    taWin64:
      Result := GetLoadConfigValues64(Index);
  end;
end;

function TJclPeImage.GetMappedAddress: NativeInt;
begin
  if StatusOK then
    Result := NativeInt(LoadedImage.MappedAddress)
  else
    Result := 0;
end;

function TJclPeImage.GetNameInStringTable(Offset: ULONG): string;
var
  Index: Integer;
begin
  Dec(Offset, SizeOf(ULONG));
  Index := 0;
  while (Offset > 0) and (Index < FStringTable.Count) do
  begin
    Dec(Offset, Length(FStringTable[Index]) + 1);
    if Offset > 0 then
      Inc(Index);
  end;

  if Offset = 0 then
    Result := FStringTable[Index]
  else
    Result := '';
end;

function TJclPeImage.GetOptionalHeader32: TImageOptionalHeader32;
begin
  if Target = taWin32 then
    Result := PImageNtHeaders32(FLoadedImage.FileHeader)^.OptionalHeader
  else
    ZeroMemory(@Result, SizeOf(Result));
end;

function TJclPeImage.GetOptionalHeader64: TImageOptionalHeader64;
begin
  if Target = taWin64 then
    Result := PImageNtHeaders64(FLoadedImage.FileHeader)^.OptionalHeader
  else
    ZeroMemory(@Result, SizeOf(Result));
end;

function TJclPeImage.GetRelocationList: TJclPeRelocList;
begin
  if FRelocationList = nil then
    FRelocationList := TJclPeRelocList.Create(Self);
  Result := FRelocationList;
end;

function TJclPeImage.GetResourceList: TJclPeRootResourceList;
begin
  if FResourceList = nil then
  begin
    FResourceVA := Directories[IMAGE_DIRECTORY_ENTRY_RESOURCE].VirtualAddress;
    if FResourceVA <> 0 then
      FResourceVA := NativeInt(RvaToVa(FResourceVA));
    FResourceList := TJclPeRootResourceList.Create(Self, nil, PImageResourceDirectory(FResourceVA));
  end;
  Result := FResourceList;
end;

function TJclPeImage.GetSectionHeader(const SectionName: string;
  out Header: PImageSectionHeader): Boolean;
var
  I: Integer;
begin
  I := FImageSections.IndexOf(SectionName);
  if I = -1 then
  begin
    Header := nil;
    Result := False;
  end
  else
  begin
    Header := PImageSectionHeader(FImageSections.Objects[I]);
    Result := True;
  end;
end;

function TJclPeImage.GetSectionName(Header: PImageSectionHeader): string;
var
  I: Integer;
begin
  I := FImageSections.IndexOfObject(TObject(Header));
  if I = -1 then
    Result := ''
  else
    Result := FImageSections[I];
end;

function TJclPeImage.GetStringTableCount: Integer;
begin
  Result := FStringTable.Count;
end;

function TJclPeImage.GetStringTableItem(Index: Integer): string;
begin
  Result := FStringTable[Index];
end;

function TJclPeImage.GetUnusedHeaderBytes: TImageDataDirectory;
begin
  CheckNotAttached;
  Result.Size := 0;
  Result.VirtualAddress := GetImageUnusedHeaderBytes(FLoadedImage, Result.Size);
  if Result.VirtualAddress = 0 then
    RaiseLastOSError;
end;

function TJclPeImage.GetVersionInfoAvailable: Boolean;
begin
  Result := StatusOK and (ResourceList.FindResource(rtVersion, '1') <> nil);
end;

class function TJclPeImage.HeaderNames(Index: TJclPeHeader): string;
begin
(*
  case Index of
    JclPeHeader_Signature:
      Result := LoadResString(@RsPeSignature);
    JclPeHeader_Machine:
      Result := LoadResString(@RsPeMachine);
    JclPeHeader_NumberOfSections:
      Result := LoadResString(@RsPeNumberOfSections);
    JclPeHeader_TimeDateStamp:
      Result := LoadResString(@RsPeTimeDateStamp);
    JclPeHeader_PointerToSymbolTable:
      Result := LoadResString(@RsPePointerToSymbolTable);
    JclPeHeader_NumberOfSymbols:
      Result := LoadResString(@RsPeNumberOfSymbols);
    JclPeHeader_SizeOfOptionalHeader:
      Result := LoadResString(@RsPeSizeOfOptionalHeader);
    JclPeHeader_Characteristics:
      Result := LoadResString(@RsPeCharacteristics);
    JclPeHeader_Magic:
      Result := LoadResString(@RsPeMagic);
    JclPeHeader_LinkerVersion:
      Result := LoadResString(@RsPeLinkerVersion);
    JclPeHeader_SizeOfCode:
      Result := LoadResString(@RsPeSizeOfCode);
    JclPeHeader_SizeOfInitializedData:
      Result := LoadResString(@RsPeSizeOfInitializedData);
    JclPeHeader_SizeOfUninitializedData:
      Result := LoadResString(@RsPeSizeOfUninitializedData);
    JclPeHeader_AddressOfEntryPoint:
      Result := LoadResString(@RsPeAddressOfEntryPoint);
    JclPeHeader_BaseOfCode:
      Result := LoadResString(@RsPeBaseOfCode);
    JclPeHeader_BaseOfData:
      Result := LoadResString(@RsPeBaseOfData);
    JclPeHeader_ImageBase:
      Result := LoadResString(@RsPeImageBase);
    JclPeHeader_SectionAlignment:
      Result := LoadResString(@RsPeSectionAlignment);
    JclPeHeader_FileAlignment:
      Result := LoadResString(@RsPeFileAlignment);
    JclPeHeader_OperatingSystemVersion:
      Result := LoadResString(@RsPeOperatingSystemVersion);
    JclPeHeader_ImageVersion:
      Result := LoadResString(@RsPeImageVersion);
    JclPeHeader_SubsystemVersion:
      Result := LoadResString(@RsPeSubsystemVersion);
    JclPeHeader_Win32VersionValue:
      Result := LoadResString(@RsPeWin32VersionValue);
    JclPeHeader_SizeOfImage:
      Result := LoadResString(@RsPeSizeOfImage);
    JclPeHeader_SizeOfHeaders:
      Result := LoadResString(@RsPeSizeOfHeaders);
    JclPeHeader_CheckSum:
      Result := LoadResString(@RsPeCheckSum);
    JclPeHeader_Subsystem:
      Result := LoadResString(@RsPeSubsystem);
    JclPeHeader_DllCharacteristics:
      Result := LoadResString(@RsPeDllCharacteristics);
    JclPeHeader_SizeOfStackReserve:
      Result := LoadResString(@RsPeSizeOfStackReserve);
    JclPeHeader_SizeOfStackCommit:
      Result := LoadResString(@RsPeSizeOfStackCommit);
    JclPeHeader_SizeOfHeapReserve:
      Result := LoadResString(@RsPeSizeOfHeapReserve);
    JclPeHeader_SizeOfHeapCommit:
      Result := LoadResString(@RsPeSizeOfHeapCommit);
    JclPeHeader_LoaderFlags:
      Result := LoadResString(@RsPeLoaderFlags);
    JclPeHeader_NumberOfRvaAndSizes:
      Result := LoadResString(@RsPeNumberOfRvaAndSizes);
  else*)
    Result := '';
  //end;
end;

function TJclPeImage.IsBrokenFormat: Boolean;
  function IsBrokenFormat32: Boolean;
  var
    OptionalHeader: TImageOptionalHeader32;
  begin
    OptionalHeader := OptionalHeader32;
    Result := not ((OptionalHeader.AddressOfEntryPoint = 0) or IsCLR);
    if Result then
    begin
      Result := (ImageSectionCount = 0);
      if not Result then
        with ImageSectionHeaders[0] do
          Result := (VirtualAddress <> OptionalHeader.BaseOfCode) or (SizeOfRawData = 0) or
            (OptionalHeader.AddressOfEntryPoint > VirtualAddress + Misc.VirtualSize) or
            (Characteristics and (IMAGE_SCN_CNT_CODE or IMAGE_SCN_MEM_WRITE) <> IMAGE_SCN_CNT_CODE);
    end;
  end;
  function IsBrokenFormat64: Boolean;
  var
    OptionalHeader: TImageOptionalHeader64;
  begin
    OptionalHeader := OptionalHeader64;
    Result := not ((OptionalHeader.AddressOfEntryPoint = 0) or IsCLR);
    if Result then
    begin
      Result := (ImageSectionCount = 0);
      if not Result then
        with ImageSectionHeaders[0] do
          Result := (VirtualAddress <> OptionalHeader.BaseOfCode) or (SizeOfRawData = 0) or
            (OptionalHeader.AddressOfEntryPoint > VirtualAddress + Misc.VirtualSize) or
            (Characteristics and (IMAGE_SCN_CNT_CODE or IMAGE_SCN_MEM_WRITE) <> IMAGE_SCN_CNT_CODE);
    end;
  end;
begin
  case Target of
    taWin32:
      Result := IsBrokenFormat32;
    taWin64:
      Result := IsBrokenFormat64;
    //taUnknown:
  else
    Result := False; // don't know how to check it
  end;
end;

function TJclPeImage.IsCLR: Boolean;
begin
  Result := DirectoryExists[IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR] and CLRHeader.HasMetadata;
end;

function TJclPeImage.IsSystemImage: Boolean;
begin
  Result := StatusOK and FLoadedImage.fSystemImage;
end;

class function TJclPeImage.LoadConfigNames(Index: TJclLoadConfig): string;
begin
  (*
  case Index of
    JclLoadConfig_Characteristics:
      Result := LoadResString(@RsPeCharacteristics);
    JclLoadConfig_TimeDateStamp:
      Result := LoadResString(@RsPeTimeDateStamp);
    JclLoadConfig_Version:
      Result := LoadResString(@RsPeVersion);
    JclLoadConfig_GlobalFlagsClear:
      Result := LoadResString(@RsPeGlobalFlagsClear);
    JclLoadConfig_GlobalFlagsSet:
      Result := LoadResString(@RsPeGlobalFlagsSet);
    JclLoadConfig_CriticalSectionDefaultTimeout:
      Result := LoadResString(@RsPeCriticalSectionDefaultTimeout);
    JclLoadConfig_DeCommitFreeBlockThreshold:
      Result := LoadResString(@RsPeDeCommitFreeBlockThreshold);
    JclLoadConfig_DeCommitTotalFreeThreshold:
      Result := LoadResString(@RsPeDeCommitTotalFreeThreshold);
    JclLoadConfig_LockPrefixTable:
      Result := LoadResString(@RsPeLockPrefixTable);
    JclLoadConfig_MaximumAllocationSize:
      Result := LoadResString(@RsPeMaximumAllocationSize);
    JclLoadConfig_VirtualMemoryThreshold:
      Result := LoadResString(@RsPeVirtualMemoryThreshold);
    JclLoadConfig_ProcessHeapFlags:
      Result := LoadResString(@RsPeProcessHeapFlags);
    JclLoadConfig_ProcessAffinityMask:
      Result := LoadResString(@RsPeProcessAffinityMask);
    JclLoadConfig_CSDVersion:
      Result := LoadResString(@RsPeCSDVersion);
    JclLoadConfig_Reserved1:
      Result := LoadResString(@RsPeReserved);
    JclLoadConfig_EditList:
      Result := LoadResString(@RsPeEditList);
    JclLoadConfig_Reserved:
      Result := LoadResString(@RsPeReserved);
  else*)
    Result := '';
  //end;
end;

procedure TJclPeImage.RaiseStatusException;
begin
  if not FNoExceptions then
    case FStatus of
      stNotPE:
        raise EJclPeImageError.Create('This is not a PE format');
      stNotFound:
        raise EJclPeImageError.CreateFmt('Cannot open file "%s"', [FFileName]);
      stNotSupported:
        raise EJclPeImageError.Create('Unknown PE target');
      stError:
        RaiseLastOSError;
    end;
end;

function TJclPeImage.RawToVa(Raw: DWORD): Pointer;
begin
  Result := Pointer(NativeInt(FLoadedImage.MappedAddress) + Raw);
end;

procedure TJclPeImage.ReadImageSections;
var
  I: Integer;
  Header: PImageSectionHeader;
  UTF8Name: AnsiString;
  SectionName: string;
begin
  if not StatusOK then
    Exit;
  Header := FLoadedImage.Sections;
  for I := 0 to FLoadedImage.NumberOfSections - 1 do
  begin
    SetLength(UTF8Name, IMAGE_SIZEOF_SHORT_NAME);
    Move(Header.Name[0], UTF8Name[1], IMAGE_SIZEOF_SHORT_NAME * SizeOf(AnsiChar));
    SectionName := Trim(string(UTF8Name));
    FImageSections.AddObject(SectionName, Pointer(Header));
    Inc(Header);
  end;
end;

procedure TJclPeImage.ReadStringTable;
var
  SymbolTable: DWORD;
  StringTablePtr: PAnsiChar;
  Ptr: PAnsiChar;
  ByteSize: ULONG;
  Start: PAnsiChar;
  StringEntry: AnsiString;
begin
  SymbolTable := LoadedImage.FileHeader.FileHeader.PointerToSymbolTable;
  if SymbolTable = 0 then
    Exit;

  StringTablePtr := PAnsiChar(LoadedImage.MappedAddress) +
                    SymbolTable +
                    (LoadedImage.FileHeader.FileHeader.NumberOfSymbols * SizeOf(IMAGE_SYMBOL));

  ByteSize := PULONG(StringTablePtr)^;
  Ptr := StringTablePtr + SizeOf(ByteSize);

  while Ptr < StringTablePtr + ByteSize do
  begin
    Start := Ptr;
    while (Ptr^ <> #0) and (Ptr < StringTablePtr + ByteSize) do
      Inc(Ptr);
    if Start <> Ptr then
    begin
      SetLength(StringEntry, Ptr - Start);
      Move(Start^, StringEntry[1], Ptr - Start);
      FStringTable.Add(string(StringEntry));
    end;
    Inc(Ptr); // to skip the #0 character
  end;
end;

function TJclPeImage.ResourceItemCreate(AEntry: PImageResourceDirectoryEntry;
  AParentItem: TJclPeResourceItem): TJclPeResourceItem;
begin
  Result := TJclPeResourceItem.Create(Self, AParentItem, AEntry);
end;

function TJclPeImage.ResourceListCreate(ADirectory: PImageResourceDirectory;
  AParentItem: TJclPeResourceItem): TJclPeResourceList;
begin
  Result := TJclPeResourceList.Create(Self, AParentItem, ADirectory);
end;

function TJclPeImage.RvaToSection(Rva: DWORD): PImageSectionHeader;
var
  I: Integer;
  SectionHeader: PImageSectionHeader;
  EndRVA: DWORD;
begin
  Result := ImageRvaToSection(FLoadedImage.FileHeader, FLoadedImage.MappedAddress, Rva);
  if Result = nil then
    for I := 0 to FImageSections.Count - 1 do
    begin
      SectionHeader := PImageSectionHeader(FImageSections.Objects[I]);
      if SectionHeader^.SizeOfRawData = 0 then
        EndRVA := SectionHeader^.Misc.VirtualSize
      else
        EndRVA := SectionHeader^.SizeOfRawData;
      Inc(EndRVA, SectionHeader^.VirtualAddress);
      if (SectionHeader^.VirtualAddress <= Rva) and (EndRVA >= Rva) then
      begin
        Result := SectionHeader;
        Break;
      end;
    end;
end;

function TJclPeImage.RvaToVa(Rva: DWORD): Pointer;
begin
  if FAttachedImage then
    Result := Pointer(NativeInt(FLoadedImage.MappedAddress) + Rva)
  else
    Result := ImageRvaToVa(FLoadedImage.FileHeader, FLoadedImage.MappedAddress, Rva, nil);
end;

function TJclPeImage.RvaToVaEx(Rva: DWORD): Pointer;
  function RvaToVaEx32(Rva: DWORD): Pointer;
  var
    OptionalHeader: TImageOptionalHeader32;
  begin
    OptionalHeader := OptionalHeader32;
    if (Rva >= OptionalHeader.ImageBase) and (Rva < (OptionalHeader.ImageBase + FLoadedImage.SizeOfImage)) then
      Dec(Rva, OptionalHeader.ImageBase);
    Result := RvaToVa(Rva);
  end;
  function RvaToVaEx64(Rva: DWORD): Pointer;
  var
    OptionalHeader: TImageOptionalHeader64;
  begin
    OptionalHeader := OptionalHeader64;
    if (Rva >= OptionalHeader.ImageBase) and (Rva < (OptionalHeader.ImageBase + FLoadedImage.SizeOfImage)) then
      Dec(Rva, OptionalHeader.ImageBase);
    Result := RvaToVa(Rva);
  end;
begin
  case Target of
    taWin32:
      Result := RvaToVaEx32(Rva);
    taWin64:
      Result := RvaToVaEx64(Rva);
    //taUnknown:
  else
    Result := nil;
  end;
end;

procedure TJclPeImage.SetFileName(const Value: TFileName);
begin
  if FFileName <> Value then
  begin
    Clear;
    FFileName := Value;
    if FFileName = '' then
      Exit;
    // OF: possible loss of data
    if MapAndLoad(PAnsiChar(AnsiString(FFileName)), nil, FLoadedImage, True, FReadOnlyAccess) then
    begin
      FTarget := PeMapImgTarget(FLoadedImage.MappedAddress);
      if FTarget <> taUnknown then
      begin
        FStatus := stOk;
        ReadImageSections;
        ReadStringTable;
        AfterOpen;
      end
      else
        FStatus := stNotSupported;
    end
    else
      case GetLastError of
        ERROR_SUCCESS:
          FStatus := stNotPE;
        ERROR_FILE_NOT_FOUND:
          FStatus := stNotFound;
      else
        FStatus := stError;
      end;
    RaiseStatusException;
  end;
end;

class function TJclPeImage.ShortSectionInfo(Characteristics: DWORD): string;
type
  TSectionCharacteristics = packed record
    Mask: DWORD;
    InfoChar: Char;
  end;
const
  Info: array [1..8] of TSectionCharacteristics = (
    (Mask: IMAGE_SCN_CNT_CODE; InfoChar: 'C'),
    (Mask: IMAGE_SCN_MEM_EXECUTE; InfoChar: 'E'),
    (Mask: IMAGE_SCN_MEM_READ; InfoChar: 'R'),
    (Mask: IMAGE_SCN_MEM_WRITE; InfoChar: 'W'),
    (Mask: IMAGE_SCN_CNT_INITIALIZED_DATA; InfoChar: 'I'),
    (Mask: IMAGE_SCN_CNT_UNINITIALIZED_DATA; InfoChar: 'U'),
    (Mask: IMAGE_SCN_MEM_SHARED; InfoChar: 'S'),
    (Mask: IMAGE_SCN_MEM_DISCARDABLE; InfoChar: 'D')
  );
var
  I: Integer;
begin
  SetLength(Result, High(Info));
  Result := '';
  for I := Low(Info) to High(Info) do
    with Info[I] do
      if (Characteristics and Mask) = Mask then
        Result := Result + InfoChar;
end;

function TJclPeImage.StatusOK: Boolean;
begin
  Result := (FStatus = stOk);
end;

class function TJclPeImage.StampToDateTime(TimeDateStamp: DWORD): TDateTime;
begin
  Result := 0;
  //Result := TimeDateStamp / SecsPerDay + UnixTimeStart
end;

procedure TJclPeImage.TryGetNamesForOrdinalImports;
begin
  if StatusOK then
  begin
    GetImportList;
    FImportList.TryGetNamesForOrdinalImports;
  end;
end;

function TJclPeImage.VerifyCheckSum: Boolean;
  function VerifyCheckSum32: Boolean;
  var
    OptionalHeader: TImageOptionalHeader32;
  begin
    OptionalHeader := OptionalHeader32;
    Result := StatusOK and ((OptionalHeader.CheckSum = 0) or (CalculateCheckSum = OptionalHeader.CheckSum));
  end;
  function VerifyCheckSum64: Boolean;
  var
    OptionalHeader: TImageOptionalHeader64;
  begin
    OptionalHeader := OptionalHeader64;
    Result := StatusOK and ((OptionalHeader.CheckSum = 0) or (CalculateCheckSum = OptionalHeader.CheckSum));
  end;
begin
  CheckNotAttached;
  case Target of
    taWin32:
      Result := VerifyCheckSum32;
    taWin64:
      Result := VerifyCheckSum64;
    //taUnknown: ;
  else
    Result := True;
  end;
end;


//=== { TJclPePackageInfo } ==================================================

constructor TJclPePackageInfo.Create(ALibHandle: THandle);
begin
  FContains := TStringList.Create;
  FRequires := TStringList.Create;
  FEnsureExtension := True;
  FSorted := True;
  ReadPackageInfo(ALibHandle);
end;

destructor TJclPePackageInfo.Destroy;
begin
  FreeAndNil(FContains);
  FreeAndNil(FRequires);
  inherited Destroy;
end;

function TJclPePackageInfo.GetContains: TStrings;
begin
  Result := FContains;
end;

function TJclPePackageInfo.GetContainsCount: Integer;
begin
  Result := Contains.Count;
end;

function TJclPePackageInfo.GetContainsFlags(Index: Integer): Byte;
begin
  Result := Byte(Contains.Objects[Index]);
end;

function TJclPePackageInfo.GetContainsNames(Index: Integer): string;
begin
  Result := Contains[Index];
end;

function TJclPePackageInfo.GetRequires: TStrings;
begin
  Result := FRequires;
end;

function TJclPePackageInfo.GetRequiresCount: Integer;
begin
  Result := Requires.Count;
end;

function TJclPePackageInfo.GetRequiresNames(Index: Integer): string;
begin
  Result := Requires[Index];
  //if FEnsureExtension then
  //  StrEnsureSuffix(BinaryExtensionPackage, Result);
end;



procedure PackageInfoProc(const Name: string; NameType: TNameType; AFlags: Byte; Param: Pointer);
begin
  with TJclPePackageInfo(Param) do
    case NameType of
      ntContainsUnit:
        Contains.AddObject(Name, Pointer(AFlags));
      ntRequiresPackage:
        Requires.Add(Name);
      ntDcpBpiName:
        SetDcpName(Name);
    end;
end;

procedure TJclPePackageInfo.ReadPackageInfo(ALibHandle: THandle);
var
  DescrResInfo: HRSRC;
  DescrResData: HGLOBAL;
begin
  FAvailable := FindResource(ALibHandle, PackageInfoResName, RT_RCDATA) <> 0;
  if FAvailable then
  begin
    GetPackageInfo(ALibHandle, Self, FFlags, PackageInfoProc);
    //if FDcpName = '' then
    //  FDcpName := PathExtractFileNameNoExt(GetModulePath(ALibHandle)) + CompilerExtensionDCP;
    if FSorted then
    begin
      FContains.Sort;
      FRequires.Sort;
    end;
  end;
  DescrResInfo := FindResource(ALibHandle, DescriptionResName, RT_RCDATA);
  if DescrResInfo <> 0 then
  begin
    DescrResData := LoadResource(ALibHandle, DescrResInfo);
    if DescrResData <> 0 then
    begin
      FDescription := WideCharLenToString(LockResource(DescrResData),
        SizeofResource(ALibHandle, DescrResInfo));
      FDescription := Trim(FDescription);
    end;
  end;
end;

procedure TJclPePackageInfo.SetDcpName(const Value: string);
begin
  FDcpName := Value;
end;

//=== { TJclPeBorForm } ======================================================

constructor TJclPeBorForm.Create(AResItem: TJclPeResourceItem;
  AFormFlags: TFilerFlags; AFormPosition: Integer;
  const AFormClassName, AFormObjectName: string);
begin
  inherited Create;
  FResItem := AResItem;
  FFormFlags := AFormFlags;
  FFormPosition := AFormPosition;
  FFormClassName := AFormClassName;
  FFormObjectName := AFormObjectName;
end;

procedure TJclPeBorForm.ConvertFormToText(const Stream: TStream);
var
  SourceStream: TJclPeResourceRawStream;
begin
  SourceStream := TJclPeResourceRawStream.Create(ResItem);
  try
    ObjectBinaryToText(SourceStream, Stream);
  finally
    SourceStream.Free;
  end;
end;

procedure TJclPeBorForm.ConvertFormToText(const Strings: TStrings);
var
  TempStream: TMemoryStream;
begin
  TempStream := TMemoryStream.Create;
  try
    ConvertFormToText(TempStream);
    TempStream.Seek(0, soFromBeginning);
    Strings.LoadFromStream(TempStream);
  finally
    TempStream.Free;
  end;
end;

function TJclPeBorForm.GetDisplayName: string;
begin
  if FFormObjectName <> '' then
    Result := FFormObjectName + ': '
  else
    Result := '';
  Result := Result + FFormClassName;
end;

//=== { TJclPeBorImage } =====================================================

constructor TJclPeBorImage.Create(ANoExceptions: Boolean);
begin
  FForms := TObjectList.Create(True);
  FPackageInfoSorted := True;
  inherited Create(ANoExceptions);
end;

destructor TJclPeBorImage.Destroy;
begin
  inherited Destroy;
  FreeAndNil(FForms);
end;

procedure TJclPeBorImage.AfterOpen;
var
  HasDVCLAL, HasPACKAGEINFO, HasPACKAGEOPTIONS: Boolean;
begin
  inherited AfterOpen;
  if StatusOK then
    with ResourceList do
    begin
      HasDVCLAL := (FindResource(rtRCData, DVclAlResName) <> nil);
      HasPACKAGEINFO := (FindResource(rtRCData, PackageInfoResName) <> nil);
      HasPACKAGEOPTIONS := (FindResource(rtRCData, PackageOptionsResName) <> nil);
      FIsPackage := HasPACKAGEINFO and HasPACKAGEOPTIONS;
      FIsBorlandImage := HasDVCLAL or FIsPackage;
    end;
end;

procedure TJclPeBorImage.Clear;
begin
  FForms.Clear;
  FreeAndNil(FPackageInfo);
  FreeLibHandle;
  inherited Clear;
  FIsBorlandImage := False;
  FIsPackage := False;
  FPackageCompilerVersion := 0;
end;

procedure TJclPeBorImage.CreateFormsList;
var
  ResTypeItem: TJclPeResourceItem;
  I: Integer;

  procedure ProcessListItem(DfmResItem: TJclPeResourceItem);
  const
    FilerSignature: array [1..4] of AnsiChar = string('TPF0');
  var
    SourceStream: TJclPeResourceRawStream;
    Reader: TReader;
    FormFlags: TFilerFlags;
    FormPosition: Integer;
    ClassName, FormName: string;
  begin
    SourceStream := TJclPeResourceRawStream.Create(DfmResItem);
    try
      if (SourceStream.Size > SizeOf(FilerSignature)) and
        (PInteger(SourceStream.Memory)^ = Integer(FilerSignature)) then
      begin
        Reader := TReader.Create(SourceStream, 4096);
        try
          Reader.ReadSignature;
          Reader.ReadPrefix(FormFlags, FormPosition);
          ClassName := Reader.ReadStr;
          FormName := Reader.ReadStr;
          FForms.Add(TJclPeBorForm.Create(DfmResItem, FormFlags, FormPosition,
            ClassName, FormName));
        finally
          Reader.Free;
        end;
      end;
    finally
      SourceStream.Free;
    end;
  end;

begin
  if StatusOK then
    with ResourceList do
    begin
      ResTypeItem := FindResource(rtRCData, '');
      if ResTypeItem <> nil then
        with ResTypeItem.List do
          for I := 0 to Count - 1 do
            ProcessListItem(Items[I].List[0]);
    end;
end;

function TJclPeBorImage.DependedPackages(List: TStrings; FullPathName, Descriptions: Boolean): Boolean;
var
  ImportList: TStringList;
  I: Integer;
  Name: string;
begin
  Result := IsBorlandImage;
  if not Result then
    Exit;
  ImportList := InternalImportedLibraries(FileName, True, FullPathName, nil);
  List.BeginUpdate;
  try
    for I := 0 to ImportList.Count - 1 do
    begin
      Name := ImportList[I];
      if Trim(ExtractFileExt(Name)) = Trim(BinaryExtensionPackage) then
      begin
        if Descriptions then
          List.Add(Name + '=' + GetPackageDescription(PChar(Name)))
        else
          List.Add(Name);
      end;
    end;
  finally
    ImportList.Free;
    List.EndUpdate;
  end;
end;

function TJclPeBorImage.FreeLibHandle: Boolean;
begin
  if FLibHandle <> 0 then
  begin
    Result := FreeLibrary(FLibHandle);
    FLibHandle := 0;
  end
  else
    Result := True;
end;

function TJclPeBorImage.GetFormCount: Integer;
begin
  if FForms.Count = 0 then
    CreateFormsList;
  Result := FForms.Count;
end;

function TJclPeBorImage.GetFormFromName(const FormClassName: string): TJclPeBorForm;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to FormCount - 1 do
    if Trim(FormClassName) = Trim(Forms[I].FormClassName) then
    begin
      Result := Forms[I];
      Break;
    end;
end;

function TJclPeBorImage.GetForms(Index: Integer): TJclPeBorForm;
begin
  Result := TJclPeBorForm(FForms[Index]);
end;

function TJclPeBorImage.GetLibHandle: THandle;
begin
  if StatusOK and (FLibHandle = 0) then
  begin
    FLibHandle := LoadLibraryEx(PChar(FileName), 0, LOAD_LIBRARY_AS_DATAFILE);
    if FLibHandle = 0 then
      RaiseLastOSError;
  end;
  Result := FLibHandle;
end;

function TJclPeBorImage.GetPackageCompilerVersion: Integer;
var
  I: Integer;
  ImportName: string;

  function CheckName: Boolean;
  begin
    Result := False;
    ImportName := AnsiUpperCase(ImportName);
    if Trim(ExtractFileExt(ImportName)) = Trim(BinaryExtensionPackage) then
    begin
      //ImportName := PathExtractFileNameNoExt(ImportName);
      //if (Length(ImportName) = 5) and
      //  CharIsDigit(ImportName[4]) and CharIsDigit(ImportName[5]) and
      //  ((Pos('RTL', ImportName) = 1) or (Pos('VCL', ImportName) = 1)) then
      //begin
      //  FPackageCompilerVersion := StrToIntDef(Copy(ImportName, 4, 2), 0);
        Result := True;
      //end;
    end;
  end;

begin
  if (FPackageCompilerVersion = 0) and IsPackage then
  begin
    with ImportList do
      for I := 0 to UniqueLibItemCount - 1 do
      begin
        ImportName := UniqueLibNames[I];
        if CheckName then
          Break;
      end;
    if FPackageCompilerVersion = 0 then
    begin
      ImportName := ExtractFileName(FileName);
      CheckName;
    end;
  end;
  Result := FPackageCompilerVersion;
end;

function TJclPeBorImage.GetPackageInfo: TJclPePackageInfo;
begin
  if StatusOK and (FPackageInfo = nil) then
  begin
    GetLibHandle;
    FPackageInfo := TJclPePackageInfo.Create(FLibHandle);
    FPackageInfo.Sorted := FPackageInfoSorted;
    FreeLibHandle;
  end;
  Result := FPackageInfo;
end;


// Mapped or loaded image related functions

function PeMapImgNtHeaders32(const BaseAddress: Pointer): PImageNtHeaders32;
begin
  Result := nil;
  if IsBadReadPtr(BaseAddress, SizeOf(TImageDosHeader)) then
    Exit;
  if (PImageDosHeader(BaseAddress)^.e_magic <> IMAGE_DOS_SIGNATURE) or
    (PImageDosHeader(BaseAddress)^._lfanew = 0) then
    Exit;
  Result := PImageNtHeaders32(NativeInt(BaseAddress) + DWORD(PImageDosHeader(BaseAddress)^._lfanew));
  if IsBadReadPtr(Result, SizeOf(TImageNtHeaders32)) or
    (Result^.Signature <> IMAGE_NT_SIGNATURE) then
      Result := nil
end;

function PeMapImgNtHeaders64(const BaseAddress: Pointer): PImageNtHeaders64;
begin
  Result := nil;
  if IsBadReadPtr(BaseAddress, SizeOf(TImageDosHeader)) then
    Exit;
  if (PImageDosHeader(BaseAddress)^.e_magic <> IMAGE_DOS_SIGNATURE) or
    (PImageDosHeader(BaseAddress)^._lfanew = 0) then
    Exit;
  Result := PImageNtHeaders64(NativeInt(BaseAddress) + DWORD(PImageDosHeader(BaseAddress)^._lfanew));
  if IsBadReadPtr(Result, SizeOf(TImageNtHeaders64)) or
    (Result^.Signature <> IMAGE_NT_SIGNATURE) then
      Result := nil
end;

function PeMapImgSize(const BaseAddress: Pointer): DWORD;
begin
  case PeMapImgTarget(BaseAddress) of
    taWin32:
      Result := PeMapImgSize32(BaseAddress);
    taWin64:
      Result := PeMapImgSize64(BaseAddress);
    //taUnknown:
  else
    Result := 0;
  end;
end;

function PeMapImgSize32(const BaseAddress: Pointer): DWORD;
var
  NtHeaders32: PImageNtHeaders32;
begin
  Result := 0;
  NtHeaders32 := PeMapImgNtHeaders32(BaseAddress);
  if Assigned(NtHeaders32) then
    Result := NtHeaders32^.OptionalHeader.SizeOfImage;
end;


function PeMapImgSize64(const BaseAddress: Pointer): DWORD;
var
  NtHeaders64: PImageNtHeaders64;
begin
  Result := 0;
  NtHeaders64 := PeMapImgNtHeaders64(BaseAddress);
  if Assigned(NtHeaders64) then
    Result := NtHeaders64^.OptionalHeader.SizeOfImage;
end;

function PeMapImgLibraryName(const BaseAddress: Pointer): string;
begin
  case PeMapImgTarget(BaseAddress) of
    taWin32:
      Result := PeMapImgLibraryName32(BaseAddress);
    taWin64:
      Result := PeMapImgLibraryName64(BaseAddress);
    //taUnknown:
  else
    Result := '';
  end;
end;

function PeMapImgLibraryName32(const BaseAddress: Pointer): string;
var
  NtHeaders: PImageNtHeaders32;
  DataDir: TImageDataDirectory;
  ExportDir: PImageExportDirectory;
  UTF8Name: AnsiString;
begin
  Result := '';
  NtHeaders := PeMapImgNtHeaders32(BaseAddress);
  if NtHeaders = nil then
    Exit;
  DataDir := NtHeaders^.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
  if DataDir.Size = 0 then
    Exit;
  ExportDir := PImageExportDirectory(NativeInt(BaseAddress) + DataDir.VirtualAddress);
  if IsBadReadPtr(ExportDir, SizeOf(TImageExportDirectory)) or (ExportDir^.Name = 0) then
    Exit;
  UTF8Name := PAnsiChar(NativeInt(BaseAddress) + ExportDir^.Name);
  Result := string(UTF8Name);
end;

function PeMapImgLibraryName64(const BaseAddress: Pointer): string;
var
  NtHeaders: PImageNtHeaders64;
  DataDir: TImageDataDirectory;
  ExportDir: PImageExportDirectory;
  UTF8Name: AnsiString;
begin
  Result := '';
  NtHeaders := PeMapImgNtHeaders64(BaseAddress);
  if NtHeaders = nil then
    Exit;
  DataDir := NtHeaders^.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
  if DataDir.Size = 0 then
    Exit;
  ExportDir := PImageExportDirectory(NativeInt(BaseAddress) + DataDir.VirtualAddress);
  if IsBadReadPtr(ExportDir, SizeOf(TImageExportDirectory)) or (ExportDir^.Name = 0) then
    Exit;
  UTF8Name := PAnsiChar(NativeInt(BaseAddress) + ExportDir^.Name);
  Result := string(UTF8Name);
end;

function PeMapImgTarget(const BaseAddress: Pointer): TJclPeTarget;
var
  ImageNtHeaders: PImageNtHeaders32;
begin
  Result := taUnknown;

  ImageNtHeaders := PeMapImgNtHeaders32(BaseAddress);
  if Assigned(ImageNtHeaders) then
    case ImageNtHeaders.FileHeader.Machine of
      IMAGE_FILE_MACHINE_I386:
        Result := taWin32;
      IMAGE_FILE_MACHINE_AMD64:
        Result := taWin64;
    end;
end;

function PeMapImgSections32(NtHeaders: PImageNtHeaders32): PImageSectionHeader;
begin
  if NtHeaders = nil then
    Result := nil
  else
    Result := PImageSectionHeader(NativeInt(@NtHeaders^.OptionalHeader) +
      NtHeaders^.FileHeader.SizeOfOptionalHeader);
end;

function PeMapImgSections32(Stream: TStream; const NtHeaders32Position: Int64; const NtHeaders32: TImageNtHeaders32;
  out ImageSectionHeaders: TImageSectionHeaderArray): Int64;
var
  SectionSize: Integer;
begin
  if NtHeaders32Position = -1 then
  begin
    SetLength(ImageSectionHeaders, 0);
    Result := -1;
  end
  else
  begin
    SetLength(ImageSectionHeaders, NtHeaders32.FileHeader.NumberOfSections);
    Result := NtHeaders32Position + SizeOf(NtHeaders32.Signature) + SizeOf(NtHeaders32.FileHeader) + NtHeaders32.FileHeader.SizeOfOptionalHeader;

    SectionSize := SizeOf(ImageSectionHeaders[0]) * Length(ImageSectionHeaders);
    if (Stream.Seek(Result, soBeginning) <> Result) or
      (Stream.Read(ImageSectionHeaders[0], SectionSize) <> SectionSize) then
      raise EJclPeImageError.CreateRes(@SReadError);
  end;
end;

function PeMapImgSections64(NtHeaders: PImageNtHeaders64): PImageSectionHeader;
begin
  if NtHeaders = nil then
    Result := nil
  else
    Result := PImageSectionHeader(NativeInt(@NtHeaders^.OptionalHeader) +
      NtHeaders^.FileHeader.SizeOfOptionalHeader);
end;

function PeMapImgSections64(Stream: TStream; const NtHeaders64Position: Int64; const NtHeaders64: TImageNtHeaders64;
  out ImageSectionHeaders: TImageSectionHeaderArray): Int64;
var
  SectionSize: Integer;
begin
  if NtHeaders64Position = -1 then
  begin
    SetLength(ImageSectionHeaders, 0);
    Result := -1;
  end
  else
  begin
    SetLength(ImageSectionHeaders, NtHeaders64.FileHeader.NumberOfSections);
    Result := NtHeaders64Position + SizeOf(NtHeaders64.Signature) + SizeOf(NtHeaders64.FileHeader) + NtHeaders64.FileHeader.SizeOfOptionalHeader;

    SectionSize := SizeOf(ImageSectionHeaders[0]) * Length(ImageSectionHeaders);
    if (Stream.Seek(Result, soBeginning) <> Result) or
      (Stream.Read(ImageSectionHeaders[0], SectionSize) <> SectionSize) then
      raise EJclPeImageError.CreateRes(@SReadError);
  end;
end;

function PeMapImgFindSection32(NtHeaders: PImageNtHeaders32;
  const SectionName: string): PImageSectionHeader;
var
  Header: PImageSectionHeader;
  I: Integer;
  P: PAnsiChar;
  UTF8Name: AnsiString;
begin
  Result := nil;
  if NtHeaders <> nil then
  begin
    UTF8Name := AnsiString(SectionName);
    P := PAnsiChar(UTF8Name);
    Header := PeMapImgSections32(NtHeaders);
    for I := 1 to NtHeaders^.FileHeader.NumberOfSections do
      if AnsiStrings.StrLComp(PAnsiChar(@Header^.Name), P, IMAGE_SIZEOF_SHORT_NAME) = 0 then
      begin
        Result := Header;
        Break;
      end
      else
        Inc(Header);
  end;
end;

function PeMapImgFindSection64(NtHeaders: PImageNtHeaders64;
  const SectionName: string): PImageSectionHeader;
var
  Header: PImageSectionHeader;
  I: Integer;
  P: PAnsiChar;
  UTF8Name: AnsiString;
begin
  Result := nil;
  if NtHeaders <> nil then
  begin
    UTF8Name := AnsiString(SectionName);
    P := PAnsiChar(UTF8Name);
    Header := PeMapImgSections64(NtHeaders);
    for I := 1 to NtHeaders^.FileHeader.NumberOfSections do
      if AnsiStrings.StrLComp(PAnsiChar(@Header^.Name), P, IMAGE_SIZEOF_SHORT_NAME) = 0 then
      begin
        Result := Header;
        Break;
      end
      else
        Inc(Header);
  end;
end;

function PeMapImgFindSectionFromModule(const BaseAddress: Pointer;
  const SectionName: string): PImageSectionHeader;
  function PeMapImgFindSectionFromModule32(const BaseAddress: Pointer;
    const SectionName: string): PImageSectionHeader;
  var
    NtHeaders32: PImageNtHeaders32;
  begin
    Result := nil;
    NtHeaders32 := PeMapImgNtHeaders32(BaseAddress);
    if Assigned(NtHeaders32) then
      Result := PeMapImgFindSection32(NtHeaders32, SectionName);
  end;
  function PeMapImgFindSectionFromModule64(const BaseAddress: Pointer;
    const SectionName: string): PImageSectionHeader;
  var
    NtHeaders64: PImageNtHeaders64;
  begin
    Result := nil;
    NtHeaders64 := PeMapImgNtHeaders64(BaseAddress);
    if Assigned(NtHeaders64) then
      Result := PeMapImgFindSection64(NtHeaders64, SectionName);
  end;
begin
  case PeMapImgTarget(BaseAddress) of
    taWin32:
      Result := PeMapImgFindSectionFromModule32(BaseAddress, SectionName);
    taWin64:
      Result := PeMapImgFindSectionFromModule64(BaseAddress, SectionName);
    //taUnknown:
  else
    Result := nil;
  end;
end;

function PeMapImgExportedVariables(const Module: HMODULE; const VariablesList: TStrings): Boolean;
var
  I: Integer;
begin
  with TJclPeImage.Create(True) do
  try
    AttachLoadedModule(Module);
    Result := StatusOK;
    if Result then
    begin
      VariablesList.BeginUpdate;
      try
        with ExportList do
          for I := 0 to Count - 1 do
            with Items[I] do
              if IsExportedVariable then
                VariablesList.AddObject(Name, MappedAddress);
      finally
        VariablesList.EndUpdate;
      end;
    end;
  finally
    Free;
  end;
end;

function ModuleFromAddr(const Addr: Pointer): HMODULE;
var
  MI: TMemoryBasicInformation;
begin
  MI.AllocationBase := nil;
  VirtualQuery(Addr, MI, SizeOf(MI));
  if MI.State <> MEM_COMMIT then
    Result := 0
  else
    Result := HMODULE(MI.AllocationBase);
end;

function PeMapImgResolvePackageThunk(Address: Pointer): Pointer;
const
  JmpInstructionCode = $25FF;
type
  PPackageThunk = ^TPackageThunk;
  TPackageThunk = packed record
    JmpInstruction: Word;
    JmpAddress: PPointer;
  end;
begin
  if not (ModuleFromAddr(Pointer(System.TObject)) <> HInstance) then
    Result := Address
  else
  if not IsBadReadPtr(Address, SizeOf(TPackageThunk)) and
    (PPackageThunk(Address)^.JmpInstruction = JmpInstructionCode) then
    Result := PPackageThunk(Address)^.JmpAddress^
  else
    Result := nil;
end;

function PeMapFindResource(const Module: HMODULE; const ResourceType: PChar;
  const ResourceName: string): Pointer;
var
  ResItem: TJclPeResourceItem;
begin
  Result := nil;
  with TJclPeImage.Create(True) do
  try
    AttachLoadedModule(Module);
    if StatusOK then
    begin
      ResItem := ResourceList.FindResource(ResourceType, PChar(ResourceName));
      if (ResItem <> nil) and ResItem.IsDirectory then
        Result := ResItem.List[0].RawEntryData;
    end;
  finally
    Free;
  end;
end;

//=== { TJclPeSectionStream } ================================================

constructor TJclPeSectionStream.Create(Instance: HMODULE; const ASectionName: string);
begin
  inherited Create;
  Initialize(Instance, ASectionName);
end;

procedure TJclPeSectionStream.Initialize(Instance: HMODULE; const ASectionName: string);
var
  Header: PImageSectionHeader;
  NtHeaders32: PImageNtHeaders32;
  NtHeaders64: PImageNtHeaders64;
  DataSize: Integer;
begin
  FInstance := Instance;
  case PeMapImgTarget(Pointer(Instance)) of
    taWin32:
      begin
        NtHeaders32 := PeMapImgNtHeaders32(Pointer(Instance));
        if NtHeaders32 = nil then
          raise EJclPeImageError.Create('This is not a PE format');
        Header := PeMapImgFindSection32(NtHeaders32, ASectionName);
      end;
    taWin64:
      begin
        NtHeaders64 := PeMapImgNtHeaders64(Pointer(Instance));
        if NtHeaders64 = nil then
          raise EJclPeImageError.Create('This is not a PE format');
        Header := PeMapImgFindSection64(NtHeaders64, ASectionName);
      end;
    //toUnknown:
  else
    raise EJclPeImageError.Create('Unknown PE target');
  end;
  if Header = nil then
    raise EJclPeImageError.CreateFmt('Section "%s" not found', [ASectionName]);
  // Borland and Microsoft seems to have swapped the meaning of this items.
  DataSize := Min(Header^.SizeOfRawData, Header^.Misc.VirtualSize);
  SetPointer(Pointer(FInstance + Header^.VirtualAddress), DataSize);
  FSectionHeader := Header^;
end;

function TJclPeSectionStream.Write(const Buffer; Count: Integer): Longint;
begin
  raise EJclPeImageError.Create('Stream is read-only');
end;


type
  TUndecorateSymbolNameA = function (DecoratedName: PAnsiChar;
    UnDecoratedName: PAnsiChar; UndecoratedLength: DWORD; Flags: DWORD): DWORD; stdcall;
// 'imagehlp.dll' 'UnDecorateSymbolName'

  TUndecorateSymbolNameW = function (DecoratedName: PWideChar;
    UnDecoratedName: PWideChar; UndecoratedLength: DWORD; Flags: DWORD): DWORD; stdcall;
// 'imagehlp.dll' 'UnDecorateSymbolNameW'

var
  UndecorateSymbolNameA: TUndecorateSymbolNameA = nil;
  UndecorateSymbolNameAFailed: Boolean = False;
  UndecorateSymbolNameW: TUndecorateSymbolNameW = nil;
  UndecorateSymbolNameWFailed: Boolean = False;

function UndecorateSymbolName(const DecoratedName: string; out UnMangled: string; Flags: DWORD): Boolean;
const
  ModuleName = 'imagehlp.dll';
  BufferSize = 512;
var
  ModuleHandle: HMODULE;
  WideBuffer: WideString;
  AnsiBuffer: AnsiString;
  Res: DWORD;
begin
  Result := False;
  if ((not Assigned(UndecorateSymbolNameA)) and (not UndecorateSymbolNameAFailed)) or
     ((not Assigned(UndecorateSymbolNameW)) and (not UndecorateSymbolNameWFailed)) then
  begin
    ModuleHandle := GetModuleHandle(ModuleName);
    if ModuleHandle = 0 then
    begin
      ModuleHandle := SafeLoadLibrary(ModuleName);
      if ModuleHandle = 0 then
        Exit;
    end;
    UndecorateSymbolNameA := GetProcAddress(ModuleHandle, 'UnDecorateSymbolName');
    UndecorateSymbolNameAFailed := not Assigned(UndecorateSymbolNameA);
    UndecorateSymbolNameW := GetProcAddress(ModuleHandle, 'UnDecorateSymbolNameW');
    UndecorateSymbolNameWFailed := not Assigned(UndecorateSymbolNameW);
  end;
  if Assigned(UndecorateSymbolNameW) then
  begin
    SetLength(WideBuffer, BufferSize);
    Res := UnDecorateSymbolNameW(PWideChar(WideString(DecoratedName)), PWideChar(WideBuffer), BufferSize, Flags);
    if Res > 0 then
    begin
      WideBuffer := Trim(WideBuffer);
      UnMangled := string(WideBuffer);
      Result := True;
    end;
  end
  else
  if Assigned(UndecorateSymbolNameA) then
  begin
    SetLength(AnsiBuffer, BufferSize);
    Res := UnDecorateSymbolNameA(PAnsiChar(AnsiString(DecoratedName)), PAnsiChar(AnsiBuffer), BufferSize, Flags);
    if Res > 0 then
    begin
      AnsiBuffer := Trim(AnsiBuffer);
      UnMangled := string(AnsiBuffer);
      Result := True;
    end;
  end;

  // For some functions UnDecorateSymbolName returns 'long'
  if Result and (UnMangled = 'long') then
    UnMangled := DecoratedName;
end;



//=== Helper assembler routines ==============================================

const
  ModuleCodeOffset = $1000;

{$STACKFRAMES OFF}

function GetFramePointer: Pointer;
asm
        {$IFDEF WIN32}
        MOV     EAX, EBP
        {$ENDIF WIN32}
        {$IFDEF WIN64}
        MOV     RAX, RBP
        {$ENDIF WIN64}
end;

function GetStackPointer: Pointer;
asm
        {$IFDEF WIN32}
        MOV     EAX, ESP
        {$ENDIF WIN32}
        {$IFDEF WIN64}
        MOV     RAX, RSP
        {$ENDIF WIN64}
end;

{$IFDEF WIN32}
function GetExceptionPointer: Pointer;
asm
        XOR     EAX, EAX
        MOV     EAX, FS:[EAX]
end;
{$ENDIF WIN32}

// Reference: Matt Pietrek, MSJ, Under the hood, on TIBs:
// http://www.microsoft.com/MSJ/archive/S2CE.HTM

function GetStackTop: NativeInt;
asm
        {$IFDEF WIN32}
        MOV     EAX, FS:[0].NT_TIB32.StackBase
        {$ENDIF WIN32}
        {$IFDEF WIN64}
        MOV     RAX, GS:[ABS 8]
        {$ENDIF WIN64}
end;

{$IFDEF STACKFRAMES_ON}
{$STACKFRAMES ON}
{$ENDIF STACKFRAMES_ON}

//=== { TJclModuleInfoList } =================================================

constructor TJclModuleInfoList.Create(ADynamicBuild, ASystemModulesOnly: Boolean);
begin
  inherited Create(True);
  FDynamicBuild := ADynamicBuild;
  FSystemModulesOnly := ASystemModulesOnly;
  if not FDynamicBuild then
    BuildModulesList;
end;

function TJclModuleInfoList.AddModule(Module: HMODULE; SystemModule: Boolean): Boolean;
begin
  Result := not IsValidModuleAddress(Pointer(Module)) and
    (CreateItemForAddress(Pointer(Module), SystemModule) <> nil);
end;

{function SortByStartAddress(Item1, Item2: Pointer): Integer;
begin
  Result := INT_PTR(TJclModuleInfo(Item2).StartAddr) - INT_PTR(TJclModuleInfo(Item1).StartAddr);
end;}

procedure TJclModuleInfoList.BuildModulesList;
var
  List: TStringList;
  I: Integer;
  CurModule: PLibModule;
begin
  if FSystemModulesOnly then
  begin
    CurModule := LibModuleList;
    while CurModule <> nil do
    begin
      CreateItemForAddress(Pointer(CurModule.Instance), True);
      CurModule := CurModule.Next;
    end;
  end
  else
  begin
    List := TStringList.Create;
    try
      LoadedModulesList(List, GetCurrentProcessId, True);
      for I := 0 to List.Count - 1 do
        CreateItemForAddress(List.Objects[I], False);
    finally
      List.Free;
    end;
  end;
  //Sort(SortByStartAddress);
end;

function TJclModuleInfoList.CreateItemForAddress(Addr: Pointer; SystemModule: Boolean): TJclModuleInfo;
var
  Module: HMODULE;
  ModuleSize: DWORD;
begin
  Result := nil;
  Module := ModuleFromAddr(Addr);
  if Module > 0 then
  begin
    ModuleSize := PeMapImgSize(Pointer(Module));
    if ModuleSize <> 0 then
    begin
      Result := TJclModuleInfo.Create;
      Result.FStartAddr := Pointer(Module);
      Result.FSize := ModuleSize;
      Result.FEndAddr := Pointer(Module + ModuleSize - 1);
      if SystemModule then
        Result.FSystemModule := True
      else
        Result.FSystemModule := IsSystemModule(Module);
    end;
  end;
  if Result <> nil then
    Add(Result);
end;

function TJclModuleInfoList.GetItems(Index: Integer): TJclModuleInfo;
begin
  Result := TJclModuleInfo(Get(Index));
end;

function TJclModuleInfoList.GetModuleFromAddress(Addr: Pointer): TJclModuleInfo;
var
  I: Integer;
  Item: TJclModuleInfo;
begin
  Result := nil;
  for I := 0 to Count - 1 do
  begin
    Item := Items[I];
    if (NativeInt(Item.StartAddr) <= NativeInt(Addr)) and (NativeInt(Item.EndAddr) > NativeInt(Addr)) then
    begin
      Result := Item;
      Break;
    end;
  end;
  if DynamicBuild and (Result = nil) then
    Result := CreateItemForAddress(Addr, False);
end;

function TJclModuleInfoList.IsSystemModuleAddress(Addr: Pointer): Boolean;
var
  Item: TJclModuleInfo;
begin
  Item := ModuleFromAddress[Addr];
  Result := (Item <> nil) and Item.SystemModule;
end;

function TJclModuleInfoList.IsValidModuleAddress(Addr: Pointer): Boolean;
begin
  Result := ModuleFromAddress[Addr] <> nil;
end;


//=== { TJclBinDebugScanner } ================================================

constructor TJclBinDebugScanner.Create(AStream: TCustomMemoryStream; CacheData: Boolean);
begin
  inherited Create;
  FCacheData := CacheData;
  FStream := AStream;
  CheckFormat;
end;

procedure TJclBinDebugScanner.CacheLineNumbers;
var
  P: Pointer;
  Value, LineNumber, C, Ln: Integer;
  CurrVA: DWORD;
begin
  if FLineNumbers = nil then
  begin
    LineNumber := 0;
    CurrVA := 0;
    C := 0;
    Ln := 0;
    P := MakePtr(PJclDbgHeader(FStream.Memory)^.LineNumbers);
    Value := 0;
    while ReadValue(P, Value) do
    begin
      Inc(CurrVA, Value);
      ReadValue(P, Value);
      Inc(LineNumber, Value);
      if C = Ln then
      begin
        if Ln < 64 then
          Ln := 64
        else
          Ln := Ln + Ln div 4;
        SetLength(FLineNumbers, Ln);
      end;
      FLineNumbers[C].VA := CurrVA;
      FLineNumbers[C].LineNumber := LineNumber;
      Inc(C);
    end;
    SetLength(FLineNumbers, C);
  end;
end;

procedure TJclBinDebugScanner.CacheProcNames;
var
  P: Pointer;
  Value, FirstWord, SecondWord, C, Ln: Integer;
  CurrAddr: DWORD;
begin
  if FProcNames = nil then
  begin
    FirstWord := 0;
    SecondWord := 0;
    CurrAddr := 0;
    C := 0;
    Ln := 0;
    P := MakePtr(PJclDbgHeader(FStream.Memory)^.Symbols);
    Value := 0;
    while ReadValue(P, Value) do
    begin
      Inc(CurrAddr, Value);
      ReadValue(P, Value);
      Inc(FirstWord, Value);
      ReadValue(P, Value);
      Inc(SecondWord, Value);
      if C = Ln then
      begin
        if Ln < 64 then
          Ln := 64
        else
          Ln := Ln + Ln div 4;
        SetLength(FProcNames, Ln);
      end;
      FProcNames[C].Addr := CurrAddr;
      FProcNames[C].FirstWord := FirstWord;
      FProcNames[C].SecondWord := SecondWord;
      Inc(C);
    end;
    SetLength(FProcNames, C);
  end;
end;

{$OVERFLOWCHECKS OFF}

procedure TJclBinDebugScanner.CheckFormat;
var
  CheckSum: Integer;
  Data, EndData: PAnsiChar;
  Header: PJclDbgHeader;
begin
  Data := FStream.Memory;
  Header := PJclDbgHeader(Data);
  FValidFormat := (Data <> nil) and (FStream.Size > SizeOf(TJclDbgHeader)) and
    (FStream.Size mod 4 = 0) and
    (Header^.Signature = JclDbgDataSignature) and (Header^.Version = JclDbgHeaderVersion);
  if FValidFormat and Header^.CheckSumValid then
  begin
    CheckSum := -Header^.CheckSum;
    EndData := Data + FStream.Size;
    while Data < EndData do
    begin
      Inc(CheckSum, PInteger(Data)^);
      Inc(PInteger(Data));
    end;
    CheckSum := (CheckSum shr 8) or (CheckSum shl 24);
    FValidFormat := (CheckSum = Header^.CheckSum);
  end;
end;

{$IFDEF OVERFLOWCHECKS_ON}
{$OVERFLOWCHECKS ON}
{$ENDIF OVERFLOWCHECKS_ON}

function SimpleCryptString(const S: AnsiString): AnsiString;
var
  I: Integer;
  C: Byte;
  P: PByte;
begin
  SetLength(Result, Length(S));
  P := PByte(Result);
  for I := 1 to Length(S) do
  begin
    C := Ord(S[I]);
    if C <> $AA then
      C := C xor $AA;
    P^ := C;
    Inc(P);
  end;
end;

function DecodeNameString(const S: PAnsiChar): string;
var
  I, B: Integer;
  C: Byte;
  P: PByte;
  Buffer: array [0..255] of AnsiChar;
begin
  Result := '';
  B := 0;
  P := PByte(S);
  case P^ of
    1:
      begin
        Inc(P);
        Result := UTF8ToString(SimpleCryptString(PAnsiChar(P)));
        Exit;
      end;
    2:
      begin
        Inc(P);
        Buffer[B] := '@';
        Inc(B);
      end;
  end;
  I := 0;
  C := 0;
  repeat
    case I and $03 of
      0:
        C := P^ and $3F;
      1:
        begin
          C := (P^ shr 6) and $03;
          Inc(P);
          Inc(C, (P^ and $0F) shl 2);
        end;
      2:
        begin
          C := (P^ shr 4) and $0F;
          Inc(P);
          Inc(C, (P^ and $03) shl 4);
        end;
      3:
        begin
          C := (P^ shr 2) and $3F;
          Inc(P);
        end;
    end;
    case C of
      $00:
        Break;
      $01..$0A:
        Inc(C, Ord('0') - $01);
      $0B..$24:
        Inc(C, Ord('A') - $0B);
      $25..$3E:
        Inc(C, Ord('a') - $25);
      $3F:
        C := Ord('_');
    end;
    Buffer[B] := AnsiChar(C);
    Inc(B);
    Inc(I);
  until B >= SizeOf(Buffer) - 1;
  Buffer[B] := #0;
  Result := UTF8ToString(Buffer);
end;

function TJclBinDebugScanner.DataToStr(A: Integer): string;
var
  P: PAnsiChar;
begin
  if A = 0 then
    Result := ''
  else
  begin
    P := PAnsiChar(NativeInt(FStream.Memory) + NativeInt(A) + NativeInt(PJclDbgHeader(FStream.Memory)^.Words) - 1);
    Result := DecodeNameString(P);
  end;
end;

function TJclBinDebugScanner.GetModuleName: string;
begin
  Result := DataToStr(PJclDbgHeader(FStream.Memory)^.ModuleName);
end;

function TJclBinDebugScanner.IsModuleNameValid(const Name: TFileName): Boolean;
begin
  Result := AnsiSameText(ModuleName, ExtractFileName(Name));
end;

function TJclBinDebugScanner.LineNumberFromAddr(Addr: DWORD): Integer;
var
  Dummy: Integer;
begin
  Result := LineNumberFromAddr(Addr, Dummy);
end;

function TJclBinDebugScanner.LineNumberFromAddr(Addr: DWORD; out Offset: Integer): Integer;
var
  P: Pointer;
  Value, LineNumber: Integer;
  CurrVA, ModuleStartVA, ItemVA: DWORD;
begin
  ModuleStartVA := ModuleStartFromAddr(Addr);
  LineNumber := 0;
  Offset := 0;
  if FCacheData then
  begin
    CacheLineNumbers;
    for Value := Length(FLineNumbers) - 1 downto 0 do
      if FLineNumbers[Value].VA <= Addr then
      begin
        if FLineNumbers[Value].VA >= ModuleStartVA then
        begin
          LineNumber := FLineNumbers[Value].LineNumber;
          Offset := Addr - FLineNumbers[Value].VA;
        end;
        Break;
      end;
  end
  else
  begin
    P := MakePtr(PJclDbgHeader(FStream.Memory)^.LineNumbers);
    CurrVA := 0;
    ItemVA := 0;
    while ReadValue(P, Value) do
    begin
      Inc(CurrVA, Value);
      if Addr < CurrVA then
      begin
        if ItemVA < ModuleStartVA then
        begin
          LineNumber := 0;
          Offset := 0;
        end;
        Break;
      end
      else
      begin
        ItemVA := CurrVA;
        ReadValue(P, Value);
        Inc(LineNumber, Value);
        Offset := Addr - CurrVA;
      end;
    end;
  end;
  Result := LineNumber;
end;

function TJclBinDebugScanner.MakePtr(A: Integer): Pointer;
begin
  Result := Pointer(NativeInt(FStream.Memory) + NativeInt(A));
end;

function TJclBinDebugScanner.ModuleNameFromAddr(Addr: DWORD): string;
var
  Value, Name: Integer;
  StartAddr: DWORD;
  P: Pointer;
begin
  P := MakePtr(PJclDbgHeader(FStream.Memory)^.Units);
  Name := 0;
  StartAddr := 0;
  Value := 0;
  while ReadValue(P, Value) do
  begin
    Inc(StartAddr, Value);
    if Addr < StartAddr then
      Break
    else
    begin
      ReadValue(P, Value);
      Inc(Name, Value);
    end;
  end;
  Result := DataToStr(Name);
end;

function TJclBinDebugScanner.ModuleStartFromAddr(Addr: DWORD): DWORD;
var
  Value: Integer;
  StartAddr, ModuleStartAddr: DWORD;
  P: Pointer;
begin
  P := MakePtr(PJclDbgHeader(FStream.Memory)^.Units);
  StartAddr := 0;
  ModuleStartAddr := DWORD(-1);
  Value := 0;
  while ReadValue(P, Value) do
  begin
    Inc(StartAddr, Value);
    if Addr < StartAddr then
      Break
    else
    begin
      ReadValue(P, Value);
      ModuleStartAddr := StartAddr;
    end;
  end;
  Result := ModuleStartAddr;
end;

function TJclBinDebugScanner.ProcNameFromAddr(Addr: DWORD): string;
var
  Dummy: Integer;
begin
  Result := ProcNameFromAddr(Addr, Dummy);
end;

function TJclBinDebugScanner.ProcNameFromAddr(Addr: DWORD; out Offset: Integer): string;
var
  P: Pointer;
  Value, FirstWord, SecondWord: Integer;
  CurrAddr, ModuleStartAddr, ItemAddr: DWORD;
begin
  ModuleStartAddr := ModuleStartFromAddr(Addr);
  FirstWord := 0;
  SecondWord := 0;
  Offset := 0;
  if FCacheData then
  begin
    CacheProcNames;
    for Value := Length(FProcNames) - 1 downto 0 do
      if FProcNames[Value].Addr <= Addr then
      begin
        if FProcNames[Value].Addr >= ModuleStartAddr then
        begin
          FirstWord := FProcNames[Value].FirstWord;
          SecondWord := FProcNames[Value].SecondWord;
          Offset := Addr - FProcNames[Value].Addr;
        end;
        Break;
      end;
  end
  else
  begin
    P := MakePtr(PJclDbgHeader(FStream.Memory)^.Symbols);
    CurrAddr := 0;
    ItemAddr := 0;
    while ReadValue(P, Value) do
    begin
      Inc(CurrAddr, Value);
      if Addr < CurrAddr then
      begin
        if ItemAddr < ModuleStartAddr then
        begin
          FirstWord := 0;
          SecondWord := 0;
          Offset := 0;
        end;
        Break;
      end
      else
      begin
        ItemAddr := CurrAddr;
        ReadValue(P, Value);
        Inc(FirstWord, Value);
        ReadValue(P, Value);
        Inc(SecondWord, Value);
        Offset := Addr - CurrAddr;
      end;
    end;
  end;
  if FirstWord <> 0 then
  begin
    Result := DataToStr(FirstWord);
    if SecondWord <> 0 then
      Result := Result + '.' + DataToStr(SecondWord)
  end
  else
    Result := '';
end;

function TJclBinDebugScanner.ReadValue(var P: Pointer; var Value: Integer): Boolean;
var
  N: Integer;
  I: Integer;
  B: Byte;
begin
  N := 0;
  I := 0;
  repeat
    B := PByte(P)^;
    Inc(PByte(P));
    Inc(N, (B and $7F) shl I);
    Inc(I, 7);
  until B and $80 = 0;
  Value := N;
  Result := (Value <> MaxInt);
end;

function TJclBinDebugScanner.SourceNameFromAddr(Addr: DWORD): string;
var
  Value, Name: Integer;
  StartAddr, ModuleStartAddr, ItemAddr: DWORD;
  P: Pointer;
  Found: Boolean;
begin
  ModuleStartAddr := ModuleStartFromAddr(Addr);
  P := MakePtr(PJclDbgHeader(FStream.Memory)^.SourceNames);
  Name := 0;
  StartAddr := 0;
  ItemAddr := 0;
  Found := False;
  Value := 0;
  while ReadValue(P, Value) do
  begin
    Inc(StartAddr, Value);
    if Addr < StartAddr then
    begin
      if ItemAddr < ModuleStartAddr then
        Name := 0
      else
        Found := True;
      Break;
    end
    else
    begin
      ItemAddr := StartAddr;
      ReadValue(P, Value);
      Inc(Name, Value);
    end;
  end;
  if Found then
    Result := DataToStr(Name)
  else
    Result := '';
end;


//=== { TJclDebugInfoSource } ================================================

constructor TJclDebugInfoSource.Create(AModule: HMODULE);
begin
  FModule := AModule;
end;

function TJclDebugInfoSource.GetFileName: TFileName;
begin
  Result := GetModulePath(FModule);
end;

function TJclDebugInfoSource.VAFromAddr(const Addr: Pointer): DWORD;
begin
  Result := DWORD(NativeUInt(Addr) - FModule - ModuleCodeOffset);
end;

//=== { TJclDebugInfoList } ==================================================

var
  DebugInfoList: TJclDebugInfoList = nil;
  InfoSourceClassList: TList = nil;
  DebugInfoCritSect: TCriticalSection;

procedure NeedDebugInfoList;
begin
  if DebugInfoList = nil then
    DebugInfoList := TJclDebugInfoList.Create;
end;

function TJclDebugInfoList.CreateDebugInfo(const Module: HMODULE): TJclDebugInfoSource;
var
  I: Integer;
begin
  NeedInfoSourceClassList;

  Result := nil;
  for I := 0 to InfoSourceClassList.Count - 1 do
  begin
    Result := TJclDebugInfoSourceClass(InfoSourceClassList.Items[I]).Create(Module);
    try
      if Result.InitializeSource then
        Break
      else
        FreeAndNil(Result);
    except
      Result.Free;
      raise;
    end;
  end;
end;

function TJclDebugInfoList.GetItemFromModule(const Module: HMODULE): TJclDebugInfoSource;
var
  I: Integer;
  TempItem: TJclDebugInfoSource;
begin
  Result := nil;
  if Module = 0 then
    Exit;
  for I := 0 to Count - 1 do
  begin
    TempItem := Items[I];
    if TempItem.Module = Module then
    begin
      Result := TempItem;
      Break;
    end;
  end;
  if Result = nil then
  begin
    Result := CreateDebugInfo(Module);
    if Result <> nil then
      Add(Result);
  end;
end;

function TJclDebugInfoList.GetItems(Index: Integer): TJclDebugInfoSource;
begin
  Result := TJclDebugInfoSource(Get(Index));
end;

function TJclDebugInfoList.GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean;
var
  Item: TJclDebugInfoSource;
begin
  FillChar(Info, SizeOf(Info), #0);
  Item := ItemFromModule[ModuleFromAddr(Addr)];
  if Item <> nil then
    Result := Item.GetLocationInfo(Addr, Info)
  else
    Result := False;
end;

class procedure TJclDebugInfoList.NeedInfoSourceClassList;
begin
  if not Assigned(InfoSourceClassList) then
  begin
    InfoSourceClassList := TList.Create;
    InfoSourceClassList.Add(Pointer(TJclDebugInfoBinary));
    InfoSourceClassList.Add(Pointer(TJclDebugInfoTD32));
    InfoSourceClassList.Add(Pointer(TJclDebugInfoSymbols));
  end;
end;

class procedure TJclDebugInfoList.RegisterDebugInfoSource(
  const InfoSourceClass: TJclDebugInfoSourceClass);
begin
  NeedInfoSourceClassList;

  InfoSourceClassList.Add(Pointer(InfoSourceClass));
end;

class procedure TJclDebugInfoList.RegisterDebugInfoSourceFirst(
  const InfoSourceClass: TJclDebugInfoSourceClass);
begin
  NeedInfoSourceClassList;

  InfoSourceClassList.Insert(0, Pointer(InfoSourceClass));
end;

class procedure TJclDebugInfoList.UnRegisterDebugInfoSource(
  const InfoSourceClass: TJclDebugInfoSourceClass);
begin
  if Assigned(InfoSourceClassList) then
    InfoSourceClassList.Remove(Pointer(InfoSourceClass));
end;


//=== { TJclDebugInfoBinary } ================================================

destructor TJclDebugInfoBinary.Destroy;
begin
  FreeAndNil(FScanner);
  FreeAndNil(FStream);
  inherited Destroy;
end;

function TJclDebugInfoBinary.GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean;
var
  VA: DWORD;
begin
  VA := VAFromAddr(Addr);
  with FScanner do
  begin
    Info.UnitName := ModuleNameFromAddr(VA);
    Result := Info.UnitName <> '';
    if Result then
    begin
      Info.Address := Addr;
      Info.ProcedureName := ProcNameFromAddr(VA, Info.OffsetFromProcName);
      Info.LineNumber := LineNumberFromAddr(VA, Info.OffsetFromLineNumber);
      Info.SourceName := SourceNameFromAddr(VA);
      Info.DebugInfo := Self;
      Info.BinaryFileName := FileName;
    end;
  end;
end;

function TJclDebugInfoBinary.InitializeSource: Boolean;
var
  JdbgFileName: TFileName;
  VerifyFileName: Boolean;
begin
  VerifyFileName := False;
  Result := (PeMapImgFindSectionFromModule(Pointer(Module), JclDbgDataResName) <> nil);
  if Result then
    FStream := TJclPeSectionStream.Create(Module, JclDbgDataResName)
  else
  begin
    JdbgFileName := ChangeFileExt(FileName, JclDbgFileExtension);
    Result := FileExists(JdbgFileName);
    if Result then
    begin
      FStream := TMemoryStream.Create();//JdbgFileName, fmOpenRead or fmShareDenyWrite);
      VerifyFileName := True;
    end;
  end;
  if Result then
  begin
    FScanner := TJclBinDebugScanner.Create(FStream, True);
    Result := FScanner.ValidFormat and
      (not VerifyFileName or FScanner.IsModuleNameValid(FileName));
  end;
end;


//=== { TJclDebugInfoTD32 } ==================================================

destructor TJclDebugInfoTD32.Destroy;
begin
  FreeAndNil(FImage);
  inherited Destroy;
end;

function TJclDebugInfoTD32.GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean;
var
  VA: DWORD;
begin
  VA := VAFromAddr(Addr);
  Info.UnitName := FImage.TD32Scanner.ModuleNameFromAddr(VA);
  Result := Info.UnitName <> '';
  if Result then
    with Info do
    begin
      Address := Addr;
      ProcedureName := FImage.TD32Scanner.ProcNameFromAddr(VA, OffsetFromProcName);
      LineNumber := FImage.TD32Scanner.LineNumberFromAddr(VA, OffsetFromLineNumber);
      SourceName := FImage.TD32Scanner.SourceNameFromAddr(VA);
      DebugInfo := Self;
      BinaryFileName := FileName;
    end;
end;

function TJclDebugInfoTD32.InitializeSource: Boolean;
begin
  FImage := TJclPeBorTD32Image.Create(True);
  try
    FImage.AttachLoadedModule(Module);
    Result := FImage.IsTD32DebugPresent;
  except
    Result := False;
  end;
end;

//=== { TJclDebugInfoSymbols } ===============================================

type
  TSymInitializeAFunc = function (hProcess: THandle; UserSearchPath: LPSTR;
    fInvadeProcess: Bool): Bool; stdcall;
  TSymInitializeWFunc = function (hProcess: THandle; UserSearchPath: LPWSTR;
    fInvadeProcess: Bool): Bool; stdcall;
  TSymGetOptionsFunc = function: DWORD; stdcall;
  TSymSetOptionsFunc = function (SymOptions: DWORD): DWORD; stdcall;
  TSymCleanupFunc = function (hProcess: THandle): Bool; stdcall;
  {$IFDEF WIN32}
  TSymGetSymFromAddrAFunc = function (hProcess: THandle; dwAddr: DWORD;
    pdwDisplacement: PDWORD; var Symbol: TImagehlpSymbolA): Bool; stdcall;
  TSymGetSymFromAddrWFunc = function (hProcess: THandle; dwAddr: DWORD;
    pdwDisplacement: PDWORD; var Symbol: TImagehlpSymbolW): Bool; stdcall;
  TSymGetModuleInfoAFunc = function (hProcess: THandle; dwAddr: DWORD;
    var ModuleInfo: TImagehlpModuleA): Bool; stdcall;
  TSymGetModuleInfoWFunc = function (hProcess: THandle; dwAddr: DWORD;
    var ModuleInfo: TImagehlpModuleW): Bool; stdcall;
  TSymLoadModuleFunc = function (hProcess: THandle; hFile: THandle; ImageName,
    ModuleName: LPSTR; BaseOfDll: DWORD; SizeOfDll: DWORD): DWORD; stdcall;
  TSymGetLineFromAddrAFunc = function (hProcess: THandle; dwAddr: DWORD;
    pdwDisplacement: PDWORD; var Line: TImageHlpLineA): Bool; stdcall;
  TSymGetLineFromAddrWFunc = function (hProcess: THandle; dwAddr: DWORD;
    pdwDisplacement: PDWORD; var Line: TImageHlpLineW): Bool; stdcall;
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  TSymGetSymFromAddrAFunc = function (hProcess: THandle; dwAddr: DWORD64;
    pdwDisplacement: PDWORD64; var Symbol: TImagehlpSymbolA64): Bool; stdcall;
  TSymGetSymFromAddrWFunc = function (hProcess: THandle; dwAddr: DWORD64;
    pdwDisplacement: PDWORD64; var Symbol: TImagehlpSymbolW64): Bool; stdcall;
  TSymGetModuleInfoAFunc = function (hProcess: THandle; dwAddr: DWORD64;
    var ModuleInfo: TImagehlpModuleA64): Bool; stdcall;
  TSymGetModuleInfoWFunc = function (hProcess: THandle; dwAddr: DWORD64;
    var ModuleInfo: TImagehlpModuleW64): Bool; stdcall;
  TSymLoadModuleFunc = function (hProcess: THandle; hFile: THandle; ImageName,
    ModuleName: LPSTR; BaseOfDll: DWORD64; SizeOfDll: DWORD): DWORD; stdcall;
  TSymGetLineFromAddrAFunc = function (hProcess: THandle; dwAddr: DWORD64;
    pdwDisplacement: PDWORD; var Line: TImageHlpLineA64): Bool; stdcall;
  TSymGetLineFromAddrWFunc = function (hProcess: THandle; dwAddr: DWORD64;
    pdwDisplacement: PDWORD; var Line: TImageHlpLineW64): Bool; stdcall;
  {$ENDIF WIN64}

var
  DebugSymbolsInitialized: Boolean = False;
  DebugSymbolsLoadFailed: Boolean = False;
  ImageHlpDllHandle: THandle = 0;
  SymInitializeAFunc: TSymInitializeAFunc = nil;
  SymInitializeWFunc: TSymInitializeWFunc = nil;
  SymGetOptionsFunc: TSymGetOptionsFunc = nil;
  SymSetOptionsFunc: TSymSetOptionsFunc = nil;
  SymCleanupFunc: TSymCleanupFunc = nil;
  SymGetSymFromAddrAFunc: TSymGetSymFromAddrAFunc = nil;
  SymGetSymFromAddrWFunc: TSymGetSymFromAddrWFunc = nil;
  SymGetModuleInfoAFunc: TSymGetModuleInfoAFunc = nil;
  SymGetModuleInfoWFunc: TSymGetModuleInfoWFunc = nil;
  SymLoadModuleFunc: TSymLoadModuleFunc = nil;
  SymGetLineFromAddrAFunc: TSymGetLineFromAddrAFunc = nil;
  SymGetLineFromAddrWFunc: TSymGetLineFromAddrWFunc = nil;

const
  ImageHlpDllName = 'imagehlp.dll';                          // do not localize
  SymInitializeAFuncName = 'SymInitialize';                  // do not localize
  SymInitializeWFuncName = 'SymInitializeW';                 // do not localize
  SymGetOptionsFuncName = 'SymGetOptions';                   // do not localize
  SymSetOptionsFuncName = 'SymSetOptions';                   // do not localize
  SymCleanupFuncName = 'SymCleanup';                         // do not localize
  {$IFDEF WIN32}
  SymGetSymFromAddrAFuncName = 'SymGetSymFromAddr';          // do not localize
  SymGetSymFromAddrWFuncName = 'SymGetSymFromAddrW';         // do not localize
  SymGetModuleInfoAFuncName = 'SymGetModuleInfo';            // do not localize
  SymGetModuleInfoWFuncName = 'SymGetModuleInfoW';           // do not localize
  SymLoadModuleFuncName = 'SymLoadModule';                   // do not localize
  SymGetLineFromAddrAFuncName = 'SymGetLineFromAddr';        // do not localize
  SymGetLineFromAddrWFuncName = 'SymGetLineFromAddrW';       // do not localize
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  SymGetSymFromAddrAFuncName = 'SymGetSymFromAddr64';        // do not localize
  SymGetSymFromAddrWFuncName = 'SymGetSymFromAddrW64';       // do not localize
  SymGetModuleInfoAFuncName = 'SymGetModuleInfo64';          // do not localize
  SymGetModuleInfoWFuncName = 'SymGetModuleInfoW64';         // do not localize
  SymLoadModuleFuncName = 'SymLoadModule64';                 // do not localize
  SymGetLineFromAddrAFuncName = 'SymGetLineFromAddr64';      // do not localize
  SymGetLineFromAddrWFuncName = 'SymGetLineFromAddrW64';     // do not localize
  {$ENDIF WIN64}


class function TJclDebugInfoSymbols.InitializeDebugSymbols: Boolean;
var
  SearchPath: string;
  SymOptions: Cardinal;
  ProcessHandle: THandle;
begin
  Result := DebugSymbolsInitialized;
  if not DebugSymbolsLoadFailed then
  begin
    Result := LoadDebugFunctions;
    DebugSymbolsLoadFailed := not Result;

    if Result then
    begin
      //if JclDebugInfoSymbolPaths <> '' then
      //begin
        {SearchPath := StrEnsureSuffix(DirSeparator, JclDebugInfoSymbolPaths);
        SearchPath := StrEnsureNoSuffix(DirSeparator, SearchPath + GetCurrentFolder);

        if GetEnvironmentVar(EnvironmentVarNtSymbolPath, EnvironmentVarValue) and (EnvironmentVarValue <> '') then
          SearchPath := StrEnsureNoSuffix(DirSeparator, StrEnsureSuffix(DirSeparator, EnvironmentVarValue) + SearchPath);
        if GetEnvironmentVar(EnvironmentVarAlternateNtSymbolPath, EnvironmentVarValue) and (EnvironmentVarValue <> '') then
          SearchPath := StrEnsureNoSuffix(DirSeparator, StrEnsureSuffix(DirSeparator, EnvironmentVarValue) + SearchPath);

        // DbgHelp.dll crashes when an empty path is specified.
        // This also means that the SearchPath must not end with a DirSeparator. }
        //SearchPath := StrRemoveEmptyPaths(SearchPath);
      //end
      //else
        // Fix crash SymLoadModuleFunc on WinXP SP3 when SearchPath=''
      //  SearchPath := GetCurrentFolder;
       SearchPath := GetCurrentDir;

      // in Windows NT, first argument is a process handle
      ProcessHandle := GetCurrentProcess;

      // Debug(WinXPSP3): SymInitializeWFunc==nil
      if Assigned(SymInitializeWFunc) then
        Result := SymInitializeWFunc(ProcessHandle, PWideChar(WideString(SearchPath)), False)
      else
      if Assigned(SymInitializeAFunc) then
        Result := SymInitializeAFunc(ProcessHandle, PAnsiChar(AnsiString(SearchPath)), False)
      else
        Result := False;

      if Result then
      begin
        SymOptions := SymGetOptionsFunc or SYMOPT_DEFERRED_LOADS
          or SYMOPT_FAIL_CRITICAL_ERRORS or SYMOPT_INCLUDE_32BIT_MODULES or SYMOPT_LOAD_LINES;
        SymOptions := SymOptions and (not (SYMOPT_NO_UNQUALIFIED_LOADS or SYMOPT_UNDNAME));
        SymSetOptionsFunc(SymOptions);
      end;

      DebugSymbolsInitialized := Result;
    end
    else
      UnloadDebugFunctions;
  end;
end;

class function TJclDebugInfoSymbols.CleanupDebugSymbols: Boolean;
begin
  Result := True;

  if DebugSymbolsInitialized then
    Result := SymCleanupFunc(GetCurrentProcess);

  UnloadDebugFunctions;
end;

function TJclDebugInfoSymbols.GetLocationInfo(const Addr: Pointer;
  out Info: TJclLocationInfo): Boolean;
const
  SymbolNameLength = 1000;
  {$IFDEF WIN32}
  SymbolSizeA = SizeOf(TImagehlpSymbolA) + SymbolNameLength * SizeOf(AnsiChar);
  SymbolSizeW = SizeOf(TImagehlpSymbolW) + SymbolNameLength * SizeOf(WideChar);
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  SymbolSizeA = SizeOf(TImagehlpSymbolA64) + SymbolNameLength * SizeOf(AnsiChar);
  SymbolSizeW = SizeOf(TImagehlpSymbolW64) + SymbolNameLength * SizeOf(WideChar);
  {$ENDIF WIN64}
var
  Displacement: DWORD;
  ProcessHandle: THandle;
  {$IFDEF WIN32}
  SymbolA: PImagehlpSymbolA;
  SymbolW: PImagehlpSymbolW;
  LineA: TImageHlpLineA;
  LineW: TImageHlpLineW;
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  SymbolA: PImagehlpSymbolA64;
  SymbolW: PImagehlpSymbolW64;
  LineA: TImageHlpLineA64;
  LineW: TImageHlpLineW64;
  {$ENDIF WIN64}
begin
  ProcessHandle := GetCurrentProcess;

  if Assigned(SymGetSymFromAddrWFunc) then
  begin
    GetMem(SymbolW, SymbolSizeW);
    try
      ZeroMemory(SymbolW, SymbolSizeW);
      SymbolW^.SizeOfStruct := SizeOf(SymbolW^);
      SymbolW^.MaxNameLength := SymbolNameLength;
      Displacement := 0;

      Result := SymGetSymFromAddrWFunc(ProcessHandle, NativeInt(Addr), @Displacement, SymbolW^);
      if Result then
      begin
        Info.DebugInfo := Self;
        Info.Address := Addr;
        Info.BinaryFileName := FileName;
        Info.OffsetFromProcName := Displacement;
        UnDecorateSymbolName(string(PWideChar(@SymbolW^.Name[0])), Info.ProcedureName, UNDNAME_NAME_ONLY or UNDNAME_NO_ARGUMENTS);
      end;
    finally
      FreeMem(SymbolW);
    end;
  end
  else
  if Assigned(SymGetSymFromAddrAFunc) then
  begin
    GetMem(SymbolA, SymbolSizeA);
    try
      ZeroMemory(SymbolA, SymbolSizeA);
      SymbolA^.SizeOfStruct := SizeOf(SymbolA^);
      SymbolA^.MaxNameLength := SymbolNameLength;
      Displacement := 0;

      Result := SymGetSymFromAddrAFunc(ProcessHandle, NativeInt(Addr), @Displacement, SymbolA^);
      if Result then
      begin
        Info.DebugInfo := Self;
        Info.Address := Addr;
        Info.BinaryFileName := FileName;
        Info.OffsetFromProcName := Displacement;
        UnDecorateSymbolName(string(PAnsiChar(@SymbolA^.Name[0])), Info.ProcedureName, UNDNAME_NAME_ONLY or UNDNAME_NO_ARGUMENTS);
      end;
    finally
      FreeMem(SymbolA);
    end;
  end
  else
    Result := False;

  // line number is optional
  if Result and Assigned(SymGetLineFromAddrWFunc) then
  begin
    ZeroMemory(@LineW, SizeOf(LineW));
    LineW.SizeOfStruct := SizeOf(LineW);
    Displacement := 0;

    if SymGetLineFromAddrWFunc(ProcessHandle, NativeInt(Addr), @Displacement, LineW) then
    begin
      Info.LineNumber := LineW.LineNumber;
      Info.UnitName := string(LineW.FileName);
      Info.OffsetFromLineNumber := Displacement;
    end;
  end
  else
  if Result and Assigned(SymGetLineFromAddrAFunc) then
  begin
    ZeroMemory(@LineA, SizeOf(LineA));
    LineA.SizeOfStruct := SizeOf(LineA);
    Displacement := 0;

    if SymGetLineFromAddrAFunc(ProcessHandle, NativeInt(Addr), @Displacement, LineA) then
    begin
      Info.LineNumber := LineA.LineNumber;
      Info.UnitName := string(LineA.FileName);
      Info.OffsetFromLineNumber := Displacement;
    end;
  end;
end;

function TJclDebugInfoSymbols.InitializeSource: Boolean;
var
  ModuleFileName: TFileName;
  {$IFDEF WIN32}
  ModuleInfoA: TImagehlpModuleA;
  ModuleInfoW: TImagehlpModuleW;
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  ModuleInfoA: TImagehlpModuleA64;
  ModuleInfoW: TImagehlpModuleW64;
  {$ENDIF WIN64}
  ProcessHandle: THandle;
begin
  Result := InitializeDebugSymbols;
  if Result then
  begin
    // in Windows NT, first argument is a process handle
    ProcessHandle := GetCurrentProcess;

    if Assigned(SymGetModuleInfoWFunc) then
    begin
      ZeroMemory(@ModuleInfoW, SizeOf(ModuleInfoW));
      ModuleInfoW.SizeOfStruct := SizeOf(ModuleInfoW);
      Result := SymGetModuleInfoWFunc(ProcessHandle, Module, ModuleInfoW);
      if not Result then
      begin
        // the symbols for this module are not loaded yet: load the module and query for the symbol again
        ModuleFileName := GetModulePath(Module);
        ZeroMemory(@ModuleInfoW, SizeOf(ModuleInfoW));
        ModuleInfoW.SizeOfStruct := SizeOf(ModuleInfoW);
        // warning: crash on WinXP SP3 when SymInitializeAFunc is called with empty SearchPath
        // OF: possible loss of data
        Result := (SymLoadModuleFunc(ProcessHandle, 0, PAnsiChar(AnsiString(ModuleFileName)), nil, 0, 0) <> 0) and
                  SymGetModuleInfoWFunc(ProcessHandle, Module, ModuleInfoW);
      end;
      Result := Result and (ModuleInfoW.BaseOfImage <> 0) and
                not (ModuleInfoW.SymType in [SymNone, SymExport]);
    end
    else
    if Assigned(SymGetModuleInfoAFunc) then
    begin
      ZeroMemory(@ModuleInfoA, SizeOf(ModuleInfoA));
      ModuleInfoA.SizeOfStruct := SizeOf(ModuleInfoA);
      Result := SymGetModuleInfoAFunc(ProcessHandle, Module, ModuleInfoA);
      if not Result then
      begin
        // the symbols for this module are not loaded yet: load the module and query for the symbol again
        ModuleFileName := GetModulePath(Module);
        ZeroMemory(@ModuleInfoA, SizeOf(ModuleInfoA));
        ModuleInfoA.SizeOfStruct := SizeOf(ModuleInfoA);
        // warning: crash on WinXP SP3 when SymInitializeAFunc is called with empty SearchPath
        // OF: possible loss of data
        Result := (SymLoadModuleFunc(ProcessHandle, 0, PAnsiChar(AnsiString(ModuleFileName)), nil, 0, 0) <> 0) and
                  SymGetModuleInfoAFunc(ProcessHandle, Module, ModuleInfoA);
      end;
      Result := Result and (ModuleInfoA.BaseOfImage <> 0) and
                not (ModuleInfoA.SymType in [SymNone, SymExport]);
    end
    else
      Result := False;
  end;
end;

class function TJclDebugInfoSymbols.LoadDebugFunctions: Boolean;
begin
  ImageHlpDllHandle := SafeLoadLibrary(ImageHlpDllName);

  if ImageHlpDllHandle <> 0 then
  begin
    SymInitializeAFunc := GetProcAddress(ImageHlpDllHandle, SymInitializeAFuncName);
    SymInitializeWFunc := GetProcAddress(ImageHlpDllHandle, SymInitializeWFuncName);
    SymGetOptionsFunc := GetProcAddress(ImageHlpDllHandle, SymGetOptionsFuncName);
    SymSetOptionsFunc := GetProcAddress(ImageHlpDllHandle, SymSetOptionsFuncName);
    SymCleanupFunc := GetProcAddress(ImageHlpDllHandle, SymCleanupFuncName);
    SymGetSymFromAddrAFunc := GetProcAddress(ImageHlpDllHandle, SymGetSymFromAddrAFuncName);
    SymGetSymFromAddrWFunc := GetProcAddress(ImageHlpDllHandle, SymGetSymFromAddrWFuncName);
    SymGetModuleInfoAFunc := GetProcAddress(ImageHlpDllHandle, SymGetModuleInfoAFuncName);
    SymGetModuleInfoWFunc := GetProcAddress(ImageHlpDllHandle, SymGetModuleInfoWFuncName);
    SymLoadModuleFunc := GetProcAddress(ImageHlpDllHandle, SymLoadModuleFuncName);
    SymGetLineFromAddrAFunc := GetProcAddress(ImageHlpDllHandle, SymGetLineFromAddrAFuncName);
    SymGetLineFromAddrWFunc := GetProcAddress(ImageHlpDllHandle, SymGetLineFromAddrWFuncName);
  end;

  // SymGetLineFromAddrFunc is optional
  Result := (ImageHlpDllHandle <> 0) and
    Assigned(SymGetOptionsFunc) and Assigned(SymSetOptionsFunc) and
    Assigned(SymCleanupFunc) and Assigned(SymLoadModuleFunc) and
    (Assigned(SymInitializeAFunc) or Assigned(SymInitializeWFunc)) and
    (Assigned(SymGetSymFromAddrAFunc) or Assigned(SymGetSymFromAddrWFunc)) and
    (Assigned(SymGetModuleInfoAFunc) or Assigned(SymGetModuleInfoWFunc));
end;

class function TJclDebugInfoSymbols.UnloadDebugFunctions: Boolean;
begin
  Result := ImageHlpDllHandle <> 0;

  if Result then
    FreeLibrary(ImageHlpDllHandle);

  ImageHlpDllHandle := 0;

  SymInitializeAFunc := nil;
  SymInitializeWFunc := nil;
  SymGetOptionsFunc := nil;
  SymSetOptionsFunc := nil;
  SymCleanupFunc := nil;
  SymGetSymFromAddrAFunc := nil;
  SymGetSymFromAddrWFunc := nil;
  SymGetModuleInfoAFunc := nil;
  SymGetModuleInfoWFunc := nil;
  SymLoadModuleFunc := nil;
  SymGetLineFromAddrAFunc := nil;
  SymGetLineFromAddrWFunc := nil;
end;

//=== Source location functions ==============================================

{$STACKFRAMES ON}

function Caller(Level: Integer; FastStackWalk: Boolean): Pointer;
var
  TopOfStack: NativeInt;
  BaseOfStack: NativeInt;
  StackFrame: PStackFrame;
begin
  Result := nil;
  try
    if FastStackWalk then
    begin
      StackFrame := GetFramePointer;
      BaseOfStack := NativeInt(StackFrame) - 1;
      TopOfStack := GetStackTop;
      while (BaseOfStack < NativeInt(StackFrame)) and (NativeInt(StackFrame) < TopOfStack) do
      begin
        if Level = 0 then
        begin
          Result := Pointer(StackFrame^.CallerAddr - 1);
          Break;
        end;
        StackFrame := PStackFrame(StackFrame^.CallerFrame);
        Dec(Level);
      end;
    end
    else
    with TJclStackInfoList.Create(False, 1, nil, False, nil, nil) do
    try
      if Level < Count then
        Result := Items[Level].CallerAddr;
    finally
      Free;
    end;
  except
    Result := nil;
  end;
end;

{$IFNDEF STACKFRAMES_ON}
{$STACKFRAMES OFF}
{$ENDIF ~STACKFRAMES_ON}

function GetLocationInfo(const Addr: Pointer): TJclLocationInfo;
begin
  try
    DebugInfoCritSect.Enter;
    try
      NeedDebugInfoList;
      DebugInfoList.GetLocationInfo(Addr, Result)
    finally
      DebugInfoCritSect.Leave;
    end;
  except
    Finalize(Result);
    FillChar(Result, SizeOf(Result), #0);
  end;
end;

function GetLocationInfo(const Addr: Pointer; out Info: TJclLocationInfo): Boolean;
begin
  try
    DebugInfoCritSect.Enter;
    try
      NeedDebugInfoList;
      Result := DebugInfoList.GetLocationInfo(Addr, Info);
    finally
      DebugInfoCritSect.Leave;
    end;
  except
    Result := False;
  end;
end;

function GetLocationInfoStr(const Addr: Pointer; IncludeModuleName, IncludeAddressOffset,
  IncludeStartProcLineOffset: Boolean; IncludeVAddress: Boolean): string;
var
  Info, StartProcInfo: TJclLocationInfo;
  OffsetStr, StartProcOffsetStr, FixedProcedureName, UnitNameWithoutUnitscope: string;
  Module : HMODULE;
begin
  OffsetStr := '';
  if GetLocationInfo(Addr, Info) then
  with Info do
  begin
    FixedProcedureName := ProcedureName;
    if Pos(UnitName + '.', FixedProcedureName) = 1 then
      FixedProcedureName := Copy(FixedProcedureName, Length(UnitName) + 2, Length(FixedProcedureName) - Length(UnitName) - 1)
    else
    if Pos('.', UnitName) > 1 then
    begin
      UnitNameWithoutUnitscope := UnitName;
      Delete(UnitNameWithoutUnitscope, 1, Pos('.', UnitNameWithoutUnitscope));
      if Pos(UnitNameWithoutUnitscope + '.', FixedProcedureName) = 1 then
        FixedProcedureName := Copy(FixedProcedureName, Length(UnitNameWithoutUnitscope) + 2, Length(FixedProcedureName) - Length(UnitNameWithoutUnitscope) - 1);
    end;

    if LineNumber > 0 then
    begin
      if IncludeStartProcLineOffset and GetLocationInfo(Pointer(NativeInt(Info.Address) -
        Cardinal(Info.OffsetFromProcName)), StartProcInfo) and (StartProcInfo.LineNumber > 0) then
          StartProcOffsetStr := Format(' + %d', [LineNumber - StartProcInfo.LineNumber])
      else
        StartProcOffsetStr := '';
      if IncludeAddressOffset then
      begin
        if OffsetFromLineNumber >= 0 then
          OffsetStr := Format(' + $%x', [OffsetFromLineNumber])
        else
          OffsetStr := Format(' - $%x', [-OffsetFromLineNumber])
      end;
      Result := Format('[%p] %s.%s (Line %u, "%s"%s)%s', [Addr, UnitName, FixedProcedureName, LineNumber,
        SourceName, StartProcOffsetStr, OffsetStr]);
    end
    else
    begin
      if IncludeAddressOffset then
        OffsetStr := Format(' + $%x', [OffsetFromProcName]);
      if UnitName <> '' then
        Result := Format('[%p] %s.%s%s', [Addr, UnitName, FixedProcedureName, OffsetStr])
      else
        Result := Format('[%p] %s%s', [Addr, FixedProcedureName, OffsetStr]);
    end;
  end
  else
  begin
    Result := Format('[%p]', [Addr]);
    IncludeVAddress := True;
  end;
  if IncludeVAddress or IncludeModuleName then
  begin
    Module := ModuleFromAddr(Addr);
    if IncludeVAddress then
    begin
      OffsetStr :=  Format('(%p) ', [Pointer(NativeUInt(Addr) - Module - ModuleCodeOffset)]);
      Result := OffsetStr + Result;
    end;
    if IncludeModuleName then
      Insert(Format('{%-12s}', [ExtractFileName(GetModulePath(Module))]), Result, 11 {$IFDEF WIN64}+8{$ENDIF});
  end;
end;

//=== { TJclStackBaseList } ==================================================

constructor TJclStackBaseList.Create;
begin
  inherited Create(True);
  FThreadID := GetCurrentThreadId;
  FTimeStamp := Now;
end;

destructor TJclStackBaseList.Destroy;
begin
  if Assigned(FOnDestroy) then
    FOnDestroy(Self);
  inherited Destroy;
end;

//=== { TJclGlobalStackList } ================================================

type
  TJclStackBaseListClass = class of TJclStackBaseList;

  TJclGlobalStackList = class(TThreadList)
  private
    FLockedTID: DWORD;
    FTIDLocked: Boolean;
    function GetExceptStackInfo(TID: DWORD): TJclStackInfoList;
    function GetLastExceptFrameList(TID: DWORD): TJclExceptFrameList;
    procedure ItemDestroyed(Sender: TObject);
  public
    destructor Destroy; override;
    procedure AddObject(AObject: TJclStackBaseList);
    procedure Clear;
    procedure LockThreadID(TID: DWORD);
    procedure UnlockThreadID;
    function FindObject(TID: DWORD; AClass: TJclStackBaseListClass): TJclStackBaseList;
    property ExceptStackInfo[TID: DWORD]: TJclStackInfoList read GetExceptStackInfo;
    property LastExceptFrameList[TID: DWORD]: TJclExceptFrameList read GetLastExceptFrameList;
  end;

var
  GlobalStackList: TJclGlobalStackList;

destructor TJclGlobalStackList.Destroy;
begin
  with LockList do
  try
    while Count > 0 do
      TObject(Items[0]).Free;
  finally
    UnlockList;
  end;
  inherited Destroy;
end;

procedure TJclGlobalStackList.AddObject(AObject: TJclStackBaseList);
var
  ReplacedObj: TObject;
begin
  AObject.FOnDestroy := ItemDestroyed;
  with LockList do
  try
    ReplacedObj := FindObject(AObject.ThreadID, TJclStackBaseListClass(AObject.ClassType));
    if ReplacedObj <> nil then
    begin
      Remove(ReplacedObj);
      ReplacedObj.Free;
    end;
    Add(AObject);
  finally
    UnlockList;
  end;
end;

procedure TJclGlobalStackList.Clear;
begin
  with LockList do
  try
    while Count > 0 do
      TObject(Items[0]).Free;
    { The following call to Clear seems to be useless, but it deallocates memory
      by setting the lists capacity back to zero. For the runtime memory leak check
      within DUnit it is important that the allocated memory before and after the
      test is equal. }
    Clear; // do not remove
  finally
    UnlockList;
  end;
end;

function TJclGlobalStackList.FindObject(TID: DWORD; AClass: TJclStackBaseListClass): TJclStackBaseList;
var
  I: Integer;
  Item: TJclStackBaseList;
begin
  Result := nil;
  with LockList do
  try
    if FTIDLocked and (GetCurrentThreadId = MainThreadID) then
      TID := FLockedTID;
    for I := 0 to Count - 1 do
    begin
      Item := Items[I];
      if (Item.ThreadID = TID) and (Item is AClass) then
      begin
        Result := Item;
        Break;
      end;
    end;
  finally
    UnlockList;
  end;
end;

function TJclGlobalStackList.GetExceptStackInfo(TID: DWORD): TJclStackInfoList;
begin
  Result := TJclStackInfoList(FindObject(TID, TJclStackInfoList));
end;

function TJclGlobalStackList.GetLastExceptFrameList(TID: DWORD): TJclExceptFrameList;
begin
  Result := TJclExceptFrameList(FindObject(TID, TJclExceptFrameList));
end;

procedure TJclGlobalStackList.ItemDestroyed(Sender: TObject);
begin
  with LockList do
  try
    Remove(Sender);
  finally
    UnlockList;
  end;
end;

procedure TJclGlobalStackList.LockThreadID(TID: DWORD);
begin
  with LockList do
  try
    if GetCurrentThreadId = MainThreadID then
    begin
      FTIDLocked := True;
      FLockedTID := TID;
    end
    else
      FTIDLocked := False;
  finally
    UnlockList;
  end;
end;

procedure TJclGlobalStackList.UnlockThreadID;
begin
  with LockList do
  try
    FTIDLocked := False;
  finally
    UnlockList;
  end;
end;

//=== { TJclGlobalModulesList } ==============================================

type
  TJclGlobalModulesList = class(TObject)
  private
    FAddedModules: TStringList;
    FLock: TCriticalSection;
    FModulesList: TJclModuleInfoList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddModule(const ModuleName: string);
    function CreateModulesList: TJclModuleInfoList;
    procedure FreeModulesList(var ModulesList: TJclModuleInfoList);
    function ValidateAddress(Addr: Pointer): Boolean;
  end;

var
  GlobalModulesList: TJclGlobalModulesList;

constructor TJclGlobalModulesList.Create;
begin
  FLock := TCriticalSection.Create;
end;

destructor TJclGlobalModulesList.Destroy;
begin
  FreeAndNil(FLock);
  FreeAndNil(FModulesList);
  FreeAndNil(FAddedModules);
  inherited Destroy;
end;

procedure TJclGlobalModulesList.AddModule(const ModuleName: string);
var
  IsMultiThreaded: Boolean;
begin
  IsMultiThreaded := IsMultiThread;
  if IsMultiThreaded then
    FLock.Enter;
  try
    if not Assigned(FAddedModules) then
    begin
      FAddedModules := TStringList.Create;
      FAddedModules.Sorted := True;
      FAddedModules.Duplicates := dupIgnore;
    end;
    FAddedModules.Add(ModuleName);
  finally
    if IsMultiThreaded then
      FLock.Leave;
  end;
end;

function TJclGlobalModulesList.CreateModulesList: TJclModuleInfoList;
var
  I: Integer;
  SystemModulesOnly: Boolean;
  IsMultiThreaded: Boolean;
  AddedModuleHandle: HMODULE;
begin
  IsMultiThreaded := IsMultiThread;
  if IsMultiThreaded then
    FLock.Enter;
  try
    if FModulesList = nil then
    begin
      SystemModulesOnly := not (stAllModules in JclStackTrackingOptions);
      Result := TJclModuleInfoList.Create(False, SystemModulesOnly);
      // Add known Borland modules collected by DLL exception hooking code
      {if SystemModulesOnly and JclHookedExceptModulesList(FHookedModules) then
        for I := Low(FHookedModules) to High(FHookedModules) do
          Result.AddModule(FHookedModules[I], True);}
      if Assigned(FAddedModules) then
        for I := 0 to FAddedModules.Count - 1 do
        begin
          AddedModuleHandle := GetModuleHandle(PChar(FAddedModules[I]));
          if (AddedModuleHandle <> 0) and
            not Assigned(Result.ModuleFromAddress[Pointer(AddedModuleHandle)]) then
            Result.AddModule(AddedModuleHandle, True);
        end;
      if stStaticModuleList in JclStackTrackingOptions then
        FModulesList := Result;
    end
    else
      Result := FModulesList;
  finally
    if IsMultiThreaded then
      FLock.Leave;
  end;
end;

procedure TJclGlobalModulesList.FreeModulesList(var ModulesList: TJclModuleInfoList);
var
  IsMultiThreaded: Boolean;
begin
  if (Self <> nil) and // happens when finalization already ran but a TJclStackInfoList is still alive
     (FModulesList <> ModulesList) then
  begin
    IsMultiThreaded := IsMultiThread;
    if IsMultiThreaded then
      FLock.Enter;
    try
      FreeAndNil(ModulesList);
    finally
      if IsMultiThreaded then
        FLock.Leave;
    end;
  end;
end;

function TJclGlobalModulesList.ValidateAddress(Addr: Pointer): Boolean;
var
  TempList: TJclModuleInfoList;
begin
  TempList := CreateModulesList;
  try
    Result := TempList.IsValidModuleAddress(Addr);
  finally
    FreeModulesList(TempList);
  end;
end;

function JclValidateModuleAddress(Addr: Pointer): Boolean;
begin
  Result := GlobalModulesList.ValidateAddress(Addr);
end;

//=== Stack info routines ====================================================

{$STACKFRAMES OFF}

function ValidCodeAddr(CodeAddr: DWORD; ModuleList: TJclModuleInfoList): Boolean;
begin
  if stAllModules in JclStackTrackingOptions then
    Result := ModuleList.IsValidModuleAddress(Pointer(CodeAddr))
  else
    Result := ModuleList.IsSystemModuleAddress(Pointer(CodeAddr));
end;

procedure CorrectExceptStackListTop(List: TJclStackInfoList; SkipFirstItem: Boolean);
begin

end;

{$STACKFRAMES ON}

function JclLastExceptStackList: TJclStackInfoList;
begin
  Result := GlobalStackList.ExceptStackInfo[GetCurrentThreadID];
end;

function JclLastExceptStackListToStrings(Strings: TStrings; IncludeModuleName, IncludeAddressOffset,
  IncludeStartProcLineOffset, IncludeVAddress: Boolean): Boolean;
var
  List: TJclStackInfoList;
begin
  List := JclLastExceptStackList;
  Result := Assigned(List);
  if Result then
    List.AddToStrings(Strings, IncludeModuleName, IncludeAddressOffset, IncludeStartProcLineOffset,
      IncludeVAddress);
end;

function JclGetExceptStackList(ThreadID: DWORD): TJclStackInfoList;
begin
  Result := GlobalStackList.ExceptStackInfo[ThreadID];
end;

function JclGetExceptStackListToStrings(ThreadID: DWORD; Strings: TStrings;
  IncludeModuleName: Boolean = False; IncludeAddressOffset: Boolean = False;
  IncludeStartProcLineOffset: Boolean = False; IncludeVAddress: Boolean = False): Boolean;
var
  List: TJclStackInfoList;
begin
  List := JclGetExceptStackList(ThreadID);
  Result := Assigned(List);
  if Result then
    List.AddToStrings(Strings, IncludeModuleName, IncludeAddressOffset, IncludeStartProcLineOffset,
      IncludeVAddress);
end;

procedure JclClearGlobalStackData;
begin
  GlobalStackList.Clear;
end;

function JclCreateStackList(Raw: Boolean; AIgnoreLevels: Integer; FirstCaller: Pointer): TJclStackInfoList;
begin
  Result := TJclStackInfoList.Create(Raw, AIgnoreLevels, FirstCaller, False, nil, nil);
  GlobalStackList.AddObject(Result);
end;

//=== { TJclStackInfoItem } ==================================================

function TJclStackInfoItem.GetCallerAddr: Pointer;
begin
  Result := Pointer(FStackInfo.CallerAddr);
end;

function TJclStackInfoItem.GetLogicalAddress: NativeInt;
begin
  Result := FStackInfo.CallerAddr - NativeInt(ModuleFromAddr(CallerAddr));
end;

//=== { TJclStackInfoList } ==================================================

constructor TJclStackInfoList.Create(ARaw: Boolean; AIgnoreLevels: Integer;
  AFirstCaller: Pointer);
begin
  Create(ARaw, AIgnoreLevels, AFirstCaller, False, nil, nil);
end;

constructor TJclStackInfoList.Create(ARaw: Boolean; AIgnoreLevels: Integer;
  AFirstCaller: Pointer; ADelayedTrace: Boolean);
begin
  Create(ARaw, AIgnoreLevels, AFirstCaller, ADelayedTrace, nil, nil);
end;

constructor TJclStackInfoList.Create(ARaw: Boolean; AIgnoreLevels: Integer;
  AFirstCaller: Pointer; ADelayedTrace: Boolean; ABaseOfStack: Pointer);
begin
  Create(ARaw, AIgnoreLevels, AFirstCaller, ADelayedTrace, ABaseOfStack, nil);
end;

constructor TJclStackInfoList.Create(ARaw: Boolean; AIgnoreLevels: Integer;
  AFirstCaller: Pointer; ADelayedTrace: Boolean; ABaseOfStack, ATopOfStack: Pointer);
var
  Item: TJclStackInfoItem;
begin
  inherited Create;
  FIgnoreLevels := AIgnoreLevels;
  FDelayedTrace := ADelayedTrace;
  FRaw := ARaw;
  BaseOfStack := NativeInt(ABaseOfStack);
  FStackOffset := 0;
  FFramePointer := ABaseOfStack;

  if ATopOfStack = nil then
    TopOfStack := GetStackTop
  else
    TopOfStack := NativeInt(ATopOfStack);

  FModuleInfoList := GlobalModulesList.CreateModulesList;
  if AFirstCaller <> nil then
  begin
    Item := TJclStackInfoItem.Create;
    Item.FStackInfo.CallerAddr := NativeInt(AFirstCaller);
    Add(Item);
  end;
  {$IFDEF WIN32}
  if DelayedTrace then
    DelayStoreStack
  else
  if Raw then
    TraceStackRaw
  else
    TraceStackFrames;
  {$ENDIF WIN32}
  {$IFDEF WIN64}
  CaptureBackTrace;
  {$ENDIF WIN64}
end;

destructor TJclStackInfoList.Destroy;
begin
  if Assigned(FStackData) then
    FreeMem(FStackData);
  GlobalModulesList.FreeModulesList(FModuleInfoList);
  inherited Destroy;
end;

{$IFDEF WIN64}
procedure TJclStackInfoList.CaptureBackTrace;
const
  InternalSkipFrames = 1; // skip this method
var
  BackTrace: array [0..127] of Pointer;
  MaxFrames: Integer;
  Hash: DWORD;
  I: Integer;
  StackInfo: TStackInfo;
  CapturedFramesCount: Word;
begin
  MaxFrames := Length(BackTrace);
  FillChar(BackTrace, SizeOf(BackTrace), #0);
  CapturedFramesCount := CaptureStackBackTrace(InternalSkipFrames, MaxFrames, @BackTrace, Hash);

  FillChar(StackInfo, SizeOf(StackInfo), #0);
  for I := 0 to CapturedFramesCount - 1 do
  begin
    StackInfo.CallerAddr := NativeInt(BackTrace[I]);
    StackInfo.Level := I;
    StoreToList(StackInfo); // skips all frames with a level less than "IgnoreLevels"
  end;
end;
{$ENDIF WIN64}

procedure TJclStackInfoList.ForceStackTracing;
begin
  if DelayedTrace and Assigned(FStackData) and not FInStackTracing then
  begin
    FInStackTracing := True;
    try
      if Raw then
        TraceStackRaw
      else
        TraceStackFrames;
      if FCorrectOnAccess then
        CorrectExceptStackListTop(Self, FSkipFirstItem);
    finally
      FInStackTracing := False;
      FDelayedTrace := False;
    end;
  end;
end;

function TJclStackInfoList.GetCount: Integer;
begin
  ForceStackTracing;
  Result := inherited Count;
end;

procedure TJclStackInfoList.AddToStrings(Strings: TStrings; IncludeModuleName, IncludeAddressOffset,
  IncludeStartProcLineOffset, IncludeVAddress: Boolean);
var
  I: Integer;
begin
  ForceStackTracing;
  Strings.BeginUpdate;
  try
    for I := 0 to Count - 1 do
      Strings.Add(GetLocationInfoStr(Items[I].CallerAddr, IncludeModuleName, IncludeAddressOffset,
        IncludeStartProcLineOffset, IncludeVAddress));
  finally
    Strings.EndUpdate;
  end;
end;

function TJclStackInfoList.GetItems(Index: Integer): TJclStackInfoItem;
begin
  ForceStackTracing;
  Result := TJclStackInfoItem(Get(Index));
end;

function TJclStackInfoList.NextStackFrame(var StackFrame: PStackFrame; var StackInfo: TStackInfo): Boolean;
var
  CallInstructionSize: Cardinal;
  StackFrameCallerFrame, NewFrame: NativeInt;
  StackFrameCallerAddr: NativeInt;
begin
  // Only report this stack frame into the StockInfo structure
  // if the StackFrame pointer, the frame pointer and the return address on the stack
  // are valid addresses
  StackFrameCallerFrame := StackInfo.CallerFrame;
  while ValidStackAddr(NativeInt(StackFrame)) do
  begin
    // CallersEBP above the previous CallersEBP
    NewFrame := StackFrame^.CallerFrame;
    if NewFrame <= StackFrameCallerFrame then
      Break;
    StackFrameCallerFrame := NewFrame;

    // CallerAddr within current process space, code segment etc.
    // CallerFrame within current thread stack. Added Mar 12 2002 per Hallvard's suggestion
    StackFrameCallerAddr := StackFrame^.CallerAddr;
    if ValidCodeAddr(StackFrameCallerAddr, FModuleInfoList) and ValidStackAddr(StackFrameCallerFrame + FStackOffset) then
    begin
      Inc(StackInfo.Level);
      StackInfo.StackFrame := StackFrame;
      StackInfo.ParamPtr := PDWORD_PTRArray(NativeInt(StackFrame) + SizeOf(TStackFrame));

      if StackFrameCallerFrame > StackInfo.CallerFrame then
        StackInfo.CallerFrame := StackFrameCallerFrame
      else
        // the frame pointer points to an address that is below
        // the last frame pointer, so it must be invalid
        Break;

      // Calculate the address of caller by subtracting the CALL instruction size (if possible)
      if ValidCallSite(StackFrameCallerAddr, CallInstructionSize) then
        StackInfo.CallerAddr := StackFrameCallerAddr - CallInstructionSize
      else
        StackInfo.CallerAddr := StackFrameCallerAddr;
      // the stack may be messed up in big projects, avoid overflow in arithmetics
      if StackFrameCallerFrame < NativeInt(StackFrame) then
        Break;
      StackInfo.DumpSize := StackFrameCallerFrame - NativeInt(StackFrame);
      StackInfo.ParamSize := (StackInfo.DumpSize - SizeOf(TStackFrame)) div 4;
      if PStackFrame(StackFrame^.CallerFrame) = StackFrame then
        Break;
      // Step to the next stack frame by following the frame pointer
      StackFrame := PStackFrame(StackFrameCallerFrame + FStackOffset);
      Result := True;
      Exit;
    end;
    // Step to the next stack frame by following the frame pointer
    StackFrame := PStackFrame(StackFrameCallerFrame + FStackOffset);
  end;
  Result := False;
end;

procedure TJclStackInfoList.StoreToList(const StackInfo: TStackInfo);
var
  Item: TJclStackInfoItem;
begin
  if ((IgnoreLevels = -1) and (StackInfo.Level > 0)) or
     (StackInfo.Level > (IgnoreLevels + 1)) then
  begin
    Item := TJclStackInfoItem.Create;
    Item.FStackInfo := StackInfo;
    Add(Item);
  end;
end;

procedure TJclStackInfoList.TraceStackFrames;
var
  StackFrame: PStackFrame;
  StackInfo: TStackInfo;
begin
  Capacity := 32; // reduce ReallocMem calls, must be > 1 because the caller's EIP register is already in the list

  // Start at level 0
  StackInfo.Level := 0;
  StackInfo.CallerFrame := 0;
  if DelayedTrace then
    // Get the current stack frame from the frame register
    StackFrame := FFramePointer
  else
  begin
    // We define the bottom of the valid stack to be the current ESP pointer
    if BaseOfStack = 0 then
      BaseOfStack := NativeInt(GetFramePointer);
    // Get a pointer to the current bottom of the stack
    StackFrame := PStackFrame(BaseOfStack);
  end;

  // We define the bottom of the valid stack to be the current frame Pointer
  // There is a TIB field called pvStackUserBase, but this includes more of the
  // stack than what would define valid stack frames.
  BaseOfStack := NativeInt(StackFrame) - 1;
  // Loop over and report all valid stackframes
  while NextStackFrame(StackFrame, StackInfo) and (inherited Count <> MaxStackTraceItems) do
    StoreToList(StackInfo);
end;

function SearchForStackPtrManipulation(StackPtr: Pointer; Proc: Pointer): Pointer;
{$IFDEF SUPPORTS_INLINE}
inline;
{$ENDIF SUPPORTS_INLINE}
{var
  Addr: PByteArray;}
begin
{  Addr := Proc;
  while (Addr <> nil) and (DWORD_PTR(Addr) > DWORD_PTR(Proc) - $100) and not IsBadReadPtr(Addr, 6) do
  begin
    if (Addr[0] = $55) and                                           // push ebp
       (Addr[1] = $8B) and (Addr[2] = $EC) then                      // mov ebp,esp
    begin
      if (Addr[3] = $83) and (Addr[4] = $C4) then                    // add esp,c8
      begin
        Result := Pointer(INT_PTR(StackPtr) - ShortInt(Addr[5]));
        Exit;
      end;
      Break;
    end;

    if (Addr[0] = $C2) and // ret $xxxx
         (((Addr[3] = $90) and (Addr[4] = $90) and (Addr[5] = $90)) or // nop
          ((Addr[3] = $CC) and (Addr[4] = $CC) and (Addr[5] = $CC))) then // int 3
      Break;

    if (Addr[0] = $C3) and // ret
         (((Addr[1] = $90) and (Addr[2] = $90) and (Addr[3] = $90)) or // nop
          ((Addr[1] = $CC) and (Addr[2] = $CC) and (Addr[3] = $CC))) then // int 3
      Break;

    if (Addr[0] = $E9) and // jmp rel-far
         (((Addr[5] = $90) and (Addr[6] = $90) and (Addr[7] = $90)) or // nop
          ((Addr[5] = $CC) and (Addr[6] = $CC) and (Addr[7] = $CC))) then // int 3
      Break;

    if (Addr[0] = $EB) and // jmp rel-near
         (((Addr[2] = $90) and (Addr[3] = $90) and (Addr[4] = $90)) or // nop
          ((Addr[2] = $CC) and (Addr[3] = $CC) and (Addr[4] = $CC))) then // int 3
      Break;

    Dec(DWORD_TR(Addr));
  end;}
  Result := StackPtr;
end;

procedure TJclStackInfoList.TraceStackRaw;
var
  StackInfo: TStackInfo;
  StackPtr: PNativeInt;
  PrevCaller: NativeInt;
  CallInstructionSize: Cardinal;
  StackTop: NativeInt;
begin
  Capacity := 32; // reduce ReallocMem calls, must be > 1 because the caller's EIP register is already in the list

  if DelayedTrace then
  begin
    if not Assigned(FStackData) then
      Exit;
    StackPtr := PNativeInt(FStackData);
  end
  else
  begin
    // We define the bottom of the valid stack to be the current ESP pointer
    if BaseOfStack = 0 then
      BaseOfStack := NativeInt(GetStackPointer);
    // Get a pointer to the current bottom of the stack
    StackPtr := PNativeInt(BaseOfStack);
  end;

  StackTop := TopOfStack;

  if Count > 0 then
    StackPtr := SearchForStackPtrManipulation(StackPtr, Pointer(Items[0].StackInfo.CallerAddr));

  // We will not be able to fill in all the fields in the StackInfo record,
  // so just blank it all out first
  FillChar(StackInfo, SizeOf(StackInfo), #0);
  // Clear the previous call address
  PrevCaller := 0;
  // Loop through all of the valid stack space
  while (NativeInt(StackPtr) < StackTop) and (inherited Count <> MaxStackTraceItems) do
  begin
    // If the current DWORD on the stack refers to a valid call site...
    if ValidCallSite(StackPtr^, CallInstructionSize) and (StackPtr^ <> PrevCaller) then
    begin
      // then pick up the callers address
      StackInfo.CallerAddr := StackPtr^ - CallInstructionSize;
      // remember to callers address so that we don't report it repeatedly
      PrevCaller := StackPtr^;
      // increase the stack level
      Inc(StackInfo.Level);
      // then report it back to our caller
      StoreToList(StackInfo);
      StackPtr := SearchForStackPtrManipulation(StackPtr, Pointer(StackInfo.CallerAddr));
    end;
    // Look at the next DWORD on the stack
    Inc(StackPtr);
  end;
  if Assigned(FStackData) then
  begin
    FreeMem(FStackData);
    FStackData := nil;
  end;
end;

{$IFDEF WIN32}
procedure TJclStackInfoList.DelayStoreStack;
var
  StackPtr: PNativeInt;
  StackDataSize: Cardinal;
begin
  if Assigned(FStackData) then
  begin
    FreeMem(FStackData);
    FStackData := nil;
  end;
  // We define the bottom of the valid stack to be the current ESP pointer
  if BaseOfStack = 0 then
  begin
    BaseOfStack := NativeInt(GetStackPointer);
    FFramePointer := GetFramePointer;
  end;

  // Get a pointer to the current bottom of the stack
  StackPtr := PNativeInt(BaseOfStack);
  if NativeInt(StackPtr) < TopOfStack then
  begin
    StackDataSize := TopOfStack - NativeInt(StackPtr);
    GetMem(FStackData, StackDataSize);
    System.Move(StackPtr^, FStackData^, StackDataSize);
    //CopyMemory(FStackData, StackPtr, StackDataSize);
  end;

  FStackOffset := Int64(FStackData) - Int64(StackPtr);
  FFramePointer := Pointer(NativeInt(FFramePointer) + FStackOffset);
  TopOfStack := TopOfStack + FStackOffset;
end;
{$ENDIF WIN32}

// Validate that the code address is a valid code site
//
// Information from Intel Manual 24319102(2).pdf, Download the 6.5 MBs from:
// http://developer.intel.com/design/pentiumii/manuals/243191.htm
// Instruction format, Chapter 2 and The CALL instruction: page 3-53, 3-54

function TJclStackInfoList.ValidCallSite(CodeAddr: NativeInt; out CallInstructionSize: Cardinal): Boolean;
var
  CodeDWORD4: DWORD;
  CodeDWORD8: DWORD;
  C4P, C8P: PDWORD;
  RM1, RM2, RM5: Byte;
begin
  // todo: 64 bit version

  // First check that the address is within range of our code segment!
  Result := CodeAddr > 8;
  if Result then
  begin
    C8P := PDWORD(CodeAddr - 8);
    C4P := PDWORD(CodeAddr - 4);
    Result := ValidCodeAddr(NativeInt(C8P), FModuleInfoList) and not IsBadReadPtr(C8P, 8);

    // Now check to see if the instruction preceding the return address
    // could be a valid CALL instruction
    if Result then
    begin
      try
        CodeDWORD8 := PDWORD(C8P)^;
        CodeDWORD4 := PDWORD(C4P)^;
        // CodeDWORD8 = (ReturnAddr-5):(ReturnAddr-6):(ReturnAddr-7):(ReturnAddr-8)
        // CodeDWORD4 = (ReturnAddr-1):(ReturnAddr-2):(ReturnAddr-3):(ReturnAddr-4)

        // ModR/M bytes contain the following bits:
        // Mod        = (76)
        // Reg/Opcode = (543)
        // R/M        = (210)
        RM1 := (CodeDWORD4 shr 24) and $7;
        RM2 := (CodeDWORD4 shr 16) and $7;
        //RM3 := (CodeDWORD4 shr 8)  and $7;
        //RM4 :=  CodeDWORD4         and $7;
        RM5 := (CodeDWORD8 shr 24) and $7;
        //RM6 := (CodeDWORD8 shr 16) and $7;
        //RM7 := (CodeDWORD8 shr 8)  and $7;

        // Check the instruction prior to the potential call site.
        // We consider it a valid call site if we find a CALL instruction there
        // Check the most common CALL variants first
        if ((CodeDWORD8 and $FF000000) = $E8000000) then
          // 5 bytes, "CALL NEAR REL32" (E8 cd)
          CallInstructionSize := 5
        else
        if ((CodeDWORD4 and $F8FF0000) = $10FF0000) and not (RM1 in [4, 5]) then
          // 2 bytes, "CALL NEAR [EAX]" (FF /2) where Reg = 010, Mod = 00, R/M <> 100 (1 extra byte)
          // and R/M <> 101 (4 extra bytes)
          CallInstructionSize := 2
        else
        if ((CodeDWORD4 and $F8FF0000) = $D0FF0000) then
          // 2 bytes, "CALL NEAR EAX" (FF /2) where Reg = 010 and Mod = 11
          CallInstructionSize := 2
        else
        if ((CodeDWORD4 and $00FFFF00) = $0014FF00) then
          // 3 bytes, "CALL NEAR [EAX+EAX*i]" (FF /2) where Reg = 010, Mod = 00 and RM = 100
          // SIB byte not validated
          CallInstructionSize := 3
        else
        if ((CodeDWORD4 and $00F8FF00) = $0050FF00) and (RM2 <> 4) then
          // 3 bytes, "CALL NEAR [EAX+$12]" (FF /2) where Reg = 010, Mod = 01 and RM <> 100 (1 extra byte)
          CallInstructionSize := 3
        else
        if ((CodeDWORD4 and $0000FFFF) = $000054FF) then
          // 4 bytes, "CALL NEAR [EAX+EAX+$12]" (FF /2) where Reg = 010, Mod = 01 and RM = 100
          // SIB byte not validated
          CallInstructionSize := 4
        else
        if ((CodeDWORD8 and $FFFF0000) = $15FF0000) then
          // 6 bytes, "CALL NEAR [$12345678]" (FF /2) where Reg = 010, Mod = 00 and RM = 101
          CallInstructionSize := 6
        else
        if ((CodeDWORD8 and $F8FF0000) = $90FF0000) and (RM5 <> 4) then
          // 6 bytes, "CALL NEAR [EAX+$12345678]" (FF /2) where Reg = 010, Mod = 10 and RM <> 100 (1 extra byte)
          CallInstructionSize := 6
        else
        if ((CodeDWORD8 and $00FFFF00) = $0094FF00) then
          // 7 bytes, "CALL NEAR [EAX+EAX+$1234567]" (FF /2) where Reg = 010, Mod = 10 and RM = 100
          CallInstructionSize := 7
        else
        if ((CodeDWORD8 and $0000FF00) = $00009A00) then
          // 7 bytes, "CALL FAR $1234:12345678" (9A ptr16:32)
          CallInstructionSize := 7
        else
          Result := False;
        // Because we're not doing a complete disassembly, we will potentially report
        // false positives. If there is odd code that uses the CALL 16:32 format, we
        // can also get false negatives.
      except
        Result := False;
      end;
    end;
  end;
end;

{$IFNDEF STACKFRAMES_ON}
{$STACKFRAMES OFF}
{$ENDIF ~STACKFRAMES_ON}

function TJclStackInfoList.ValidStackAddr(StackAddr: NativeInt): Boolean;
begin
  Result := (BaseOfStack < StackAddr) and (StackAddr < TopOfStack);
end;

{$OVERFLOWCHECKS OFF}

function GetJmpDest(Jmp: PJmpInstruction): Pointer;
begin
  // TODO : 64 bit version
  if Jmp^.opCode = $E9 then
    Result := Pointer(NativeInt(Jmp) + NativeInt(Jmp^.distance) + 5)
  else
  if Jmp.opCode = $EB then
    Result := Pointer(NativeInt(Jmp) + NativeInt(ShortInt(Jmp^.distance)) + 2)
  else
    Result := nil;
  if (Result <> nil) and (PJmpTable(Result).OPCode = $25FF) then
    if not IsBadReadPtr(PJmpTable(Result).Ptr, SizeOf(Pointer)) then
      Result := Pointer(PNativeInt(PJmpTable(Result).Ptr)^);
end;

{$IFDEF OVERFLOWCHECKS_ON}
{$OVERFLOWCHECKS ON}
{$ENDIF OVERFLOWCHECKS_ON}

//=== { TJclExceptFrame } ====================================================

constructor TJclExceptFrame.Create(AFrameLocation: Pointer; AExcDesc: PExcDesc);
begin
  inherited Create;
  FFrameKind := efkUnknown;
  FFrameLocation := AFrameLocation;
  FCodeLocation := nil;
  AnalyseExceptFrame(AExcDesc);
end;

{$RANGECHECKS OFF}

procedure TJclExceptFrame.AnalyseExceptFrame(AExcDesc: PExcDesc);
var
  Dest: Pointer;
  LocInfo: TJclLocationInfo;
  FixedProcedureName: string;
  DotPos, I: Integer;
begin
  Dest := GetJmpDest(@AExcDesc^.Jmp);
  if Dest <> nil then
  begin
    // get frame kind
    LocInfo := GetLocationInfo(Dest);
    if CompareText(LocInfo.UnitName, 'system') = 0 then
    begin
      FixedProcedureName := LocInfo.ProcedureName;
      DotPos := Pos('.', FixedProcedureName);
      if DotPos > 0 then
        FixedProcedureName := Copy(FixedProcedureName, DotPos + 1, Length(FixedProcedureName) - DotPos);
      if CompareText(FixedProcedureName, '@HandleAnyException') = 0 then
        FFrameKind := efkAnyException
      else
      if CompareText(FixedProcedureName, '@HandleOnException') = 0 then
        FFrameKind := efkOnException
      else
      if CompareText(FixedProcedureName, '@HandleAutoException') = 0 then
        FFrameKind := efkAutoException
      else
      if CompareText(FixedProcedureName, '@HandleFinally') = 0 then
        FFrameKind := efkFinally;
    end;

    // get location
    if FFrameKind <> efkUnknown then
    begin
      FCodeLocation := GetJmpDest(PJmpInstruction(NativeInt(@AExcDesc^.Instructions)));
      if FCodeLocation = nil then
        FCodeLocation := @AExcDesc^.Instructions;
    end
    else
    begin
      FCodeLocation := GetJmpDest(PJmpInstruction(NativeInt(AExcDesc)));
      if FCodeLocation = nil then
        FCodeLocation := AExcDesc;
    end;

    // get on handlers
    if FFrameKind = efkOnException then
    begin
      SetLength(FExcTab, AExcDesc^.Cnt);
      for I := 0 to AExcDesc^.Cnt - 1 do
      begin
        if AExcDesc^.ExcTab[I].VTable = nil then
        begin
          SetLength(FExcTab, I);
          Break;
        end
        else
          FExcTab[I] := AExcDesc^.ExcTab[I];
      end;
    end;
  end;
end;

{$IFDEF RANGECHECKS_ON}
{$RANGECHECKS ON}
{$ENDIF RANGECHECKS_ON}

function TJclExceptFrame.Handles(ExceptObj: TObject): Boolean;
var
  Handler: Pointer;
begin
  Result := HandlerInfo(ExceptObj, Handler);
end;

{$OVERFLOWCHECKS OFF}

function TJclExceptFrame.HandlerInfo(ExceptObj: TObject; out HandlerAt: Pointer): Boolean;
var
  I: Integer;
  ObjVTable, VTable, ParentVTable: Pointer;
begin
  Result := FrameKind in [efkAnyException, efkAutoException];
  if not Result and (FrameKind = efkOnException) then
  begin
    HandlerAt := nil;
    ObjVTable := Pointer(ExceptObj.ClassType);
    for I := Low(FExcTab) to High(FExcTab) do
    begin
      VTable := ObjVTable;
      Result := FExcTab[I].VTable = nil;
      while (not Result) and (VTable <> nil) do
      begin
        Result := (FExcTab[I].VTable = VTable) or
          (PShortString(PPointer(PNativeInt(FExcTab[I].VTable)^ + NativeInt(vmtClassName))^)^ =
           PShortString(PPointer(NativeInt(VTable) + NativeInt(vmtClassName))^)^);
        if Result then
          HandlerAt := FExcTab[I].Handler
        else
        begin
          ParentVTable := TClass(VTable).ClassParent;
          if ParentVTable = VTable then
            VTable := nil
          else
            VTable := ParentVTable;
        end;
      end;
      if Result then
        Break;
    end;
  end
  else
  if Result then
    HandlerAt := FCodeLocation
  else
    HandlerAt := nil;
end;

{$IFDEF OVERFLOWCHECKS_ON}
{$OVERFLOWCHECKS ON}
{$ENDIF OVERFLOWCHECKS_ON}

//=== { TJclExceptFrameList } ================================================

constructor TJclExceptFrameList.Create(AIgnoreLevels: Integer);
begin
  inherited Create;
  FIgnoreLevels := AIgnoreLevels;
  TraceExceptionFrames;
end;

function TJclExceptFrameList.AddFrame(AFrame: PExcFrame): TJclExceptFrame;
begin
  Result := TJclExceptFrame.Create(AFrame, AFrame^.Desc);
  Add(Result);
end;

function TJclExceptFrameList.GetItems(Index: Integer): TJclExceptFrame;
begin
  Result := TJclExceptFrame(Get(Index));
end;

procedure TJclExceptFrameList.TraceExceptionFrames;
{$IFDEF WIN32}
var
  ExceptionPointer: PExcFrame;
  Level: Integer;
  ModulesList: TJclModuleInfoList;
begin
  Clear;
  ModulesList := GlobalModulesList.CreateModulesList;
  try
    Level := 0;
    ExceptionPointer := GetExceptionPointer;
    while NativeInt(ExceptionPointer) <> High(NativeInt) do
    begin
      if (Level >= IgnoreLevels) and ValidCodeAddr(NativeInt(ExceptionPointer^.Desc), ModulesList) then
        AddFrame(ExceptionPointer);
      Inc(Level);
      ExceptionPointer := ExceptionPointer^.next;
    end;
  finally
    GlobalModulesList.FreeModulesList(ModulesList);
  end;
end;
{$ENDIF WIN32}
{$IFDEF WIN64}
begin
  // TODO: 64-bit version
end;
{$ENDIF WIN64}

//=== Exception hooking ======================================================

var
  IgnoredExceptions: TThreadList = nil;
  IgnoredExceptionClassNames: TStringList = nil;

procedure AddIgnoredException(const ExceptionClass: TClass);
begin
  if Assigned(ExceptionClass) then
  begin
    if not Assigned(IgnoredExceptions) then
      IgnoredExceptions := TThreadList.Create;

    IgnoredExceptions.Add(ExceptionClass);
  end;
end;

function GetExceptionStackInfo(P: PExceptionRecord): Pointer;
const
  cDelphiException = $0EEDFADE;
var
  Stack: TJclStackInfoList;
  Str: TStringList;
  Trace: String;
  Sz: Integer;
begin
  if P^.ExceptionCode = cDelphiException then
    Stack := JclCreateStackList(False, 3, P^.ExceptAddr)
  else
    Stack := JclCreateStackList(False, 3, P^.ExceptionAddress);
  try
    Str := TStringList.Create;
    try
      Stack.AddToStrings(Str, True);//, True, True, True);
      Trace := Str.Text;
    finally
      FreeAndNil(Str);
    end;
  finally
    FreeAndNil(Stack);
  end;

  if Trace <> '' then
  begin
    Sz := (Length(Trace) + 1) * SizeOf(Char);
    GetMem(Result, Sz);
    Move(Pointer(Trace)^, Result^, Sz);
  end
  else
    Result := nil;
end;

function GetStackInfoString(Info: Pointer): string;
begin
  Result := PChar(Info);
end;

procedure CleanUpStackInfo(Info: Pointer);
begin
  FreeMem(Info);
end;

procedure SetupExceptionProcs;
begin
  if not Assigned(Exception.GetExceptionStackInfoProc) then
  begin
    Exception.GetExceptionStackInfoProc := GetExceptionStackInfo;
    Exception.GetStackInfoStringProc := GetStackInfoString;
    Exception.CleanUpStackInfoProc := CleanUpStackInfo;
  end;
end;

procedure ResetExceptionProcs;
begin
  if @Exception.GetExceptionStackInfoProc = @GetExceptionStackInfo then
  begin
    Exception.GetExceptionStackInfoProc := nil;
    Exception.GetStackInfoStringProc := nil;
    Exception.CleanUpStackInfoProc := nil;
  end;
end;

procedure GetProcedureAddress(var P: Pointer; const ModuleName, ProcName: string);
var
  ModuleHandle: HMODULE;
begin
  if not Assigned(P) then
  begin
    ModuleHandle := GetModuleHandle(PChar(ModuleName));
    if ModuleHandle = 0 then
    begin
      ModuleHandle := SafeLoadLibrary(PChar(ModuleName));
      if ModuleHandle = 0 then
        raise Exception.CreateFmt('Library not found: %s', [ModuleName]);
    end;
    P := GetProcAddress(ModuleHandle, PChar(ProcName));
    if not Assigned(P) then
      raise Exception.CreateFmt('Function not found: %s.%s', [ModuleName, ProcName]);
  end;
end;

const
  ImageHlpLib = 'imagehlp.dll';

type
  TReBaseImage = function (CurrentImageName: PAnsiChar; SymbolPath: PAnsiChar; fReBase: BOOL;
    fRebaseSysfileOk: BOOL; fGoingDown: BOOL; CheckImageSize: ULONG;
    var OldImageSize: LongWord; var OldImageBase: NativeInt;
    var NewImageSize: LongWord; var NewImageBase: NativeInt; TimeStamp: ULONG): BOOL; stdcall;

var
  _ReBaseImage: TReBaseImage = nil;

function ReBaseImage(CurrentImageName: PAnsiChar; SymbolPath: PAnsiChar; fReBase: BOOL;
  fRebaseSysfileOk: BOOL; fGoingDown: BOOL; CheckImageSize: ULONG;
  var OldImageSize: LongWord; var OldImageBase: NativeInt;
  var NewImageSize: LongWord; var NewImageBase: NativeInt; TimeStamp: ULONG): BOOL;
begin
  GetProcedureAddress(Pointer(@_ReBaseImage), ImageHlpLib, 'ReBaseImage');
  Result := _ReBaseImage(CurrentImageName, SymbolPath, fReBase, fRebaseSysfileOk, fGoingDown, CheckImageSize, OldImageSize, OldImageBase, NewImageSize, NewImageBase, TimeStamp);
end;

type
  TReBaseImage64 = function (CurrentImageName: PAnsiChar; SymbolPath: PAnsiChar; fReBase: BOOL;
    fRebaseSysfileOk: BOOL; fGoingDown: BOOL; CheckImageSize: ULONG;
    var OldImageSize: LongWord; var OldImageBase: Int64;
    var NewImageSize: LongWord; var NewImageBase: Int64; TimeStamp: ULONG): BOOL; stdcall;

var
  _ReBaseImage64: TReBaseImage64 = nil;

function ReBaseImage64(CurrentImageName: PAnsiChar; SymbolPath: PAnsiChar; fReBase: BOOL;
  fRebaseSysfileOk: BOOL; fGoingDown: BOOL; CheckImageSize: ULONG;
  var OldImageSize: LongWord; var OldImageBase: Int64;
  var NewImageSize: LongWord; var NewImageBase: Int64; TimeStamp: ULONG): BOOL;
begin
  GetProcedureAddress(Pointer(@_ReBaseImage64), ImageHlpLib, 'ReBaseImage64');
  Result := _ReBaseImage64(CurrentImageName, SymbolPath, fReBase, fRebaseSysfileOk, fGoingDown, CheckImageSize, OldImageSize, OldImageBase, NewImageSize, NewImageBase, TimeStamp);
end;

type
  TCheckSumMappedFile = function (BaseAddress: Pointer; FileLength: DWORD;
    out HeaderSum, CheckSum: DWORD): PImageNtHeaders; stdcall;

var
  _CheckSumMappedFile: TCheckSumMappedFile = nil;

function CheckSumMappedFile(BaseAddress: Pointer; FileLength: DWORD;
  out HeaderSum, CheckSum: DWORD): PImageNtHeaders;
begin
  GetProcedureAddress(Pointer(@_CheckSumMappedFile), ImageHlpLib, 'CheckSumMappedFile');
  Result := _CheckSumMappedFile(BaseAddress, FileLength, HeaderSum, CheckSum);
end;

type
  TGetImageUnusedHeaderBytes = function (const LoadedImage: LOADED_IMAGE;
    var SizeUnusedHeaderBytes: DWORD): DWORD; stdcall;

var
  _GetImageUnusedHeaderBytes: TGetImageUnusedHeaderBytes = nil;

function GetImageUnusedHeaderBytes(const LoadedImage: LOADED_IMAGE;
  var SizeUnusedHeaderBytes: DWORD): DWORD;
begin
  GetProcedureAddress(Pointer(@_GetImageUnusedHeaderBytes), ImageHlpLib, 'GetImageUnusedHeaderBytes');
  Result := _GetImageUnusedHeaderBytes(LoadedImage, SizeUnusedHeaderBytes);
end;

type
  TMapAndLoad = function (ImageName, DllPath: PAnsiChar; var LoadedImage: LOADED_IMAGE;
    DotDll: BOOL; ReadOnly: BOOL): BOOL; stdcall;

var
  _MapAndLoad: TMapAndLoad = nil;

function MapAndLoad(ImageName, DllPath: PAnsiChar; var LoadedImage: LOADED_IMAGE;
  DotDll: BOOL; ReadOnly: BOOL): BOOL;
begin
  GetProcedureAddress(Pointer(@_MapAndLoad), ImageHlpLib, 'MapAndLoad');
  Result := _MapAndLoad(ImageName, DllPath, LoadedImage, DotDll, ReadOnly);
end;

type
  TUnMapAndLoad = function (const LoadedImage: LOADED_IMAGE): BOOL; stdcall;

var
  _UnMapAndLoad: TUnMapAndLoad = nil;

function UnMapAndLoad(const LoadedImage: LOADED_IMAGE): BOOL;
begin
  GetProcedureAddress(Pointer(@_UnMapAndLoad), ImageHlpLib, 'UnMapAndLoad');
  Result := _UnMapAndLoad(LoadedImage);
end;

type
  TTouchFileTimes = function (const FileHandle: THandle; const pSystemTime: TSystemTime): BOOL; stdcall;

var
  _TouchFileTimes: TTouchFileTimes = nil;

function TouchFileTimes(const FileHandle: THandle; const pSystemTime: TSystemTime): BOOL;
begin
  GetProcedureAddress(Pointer(@_TouchFileTimes), ImageHlpLib, 'TouchFileTimes');
  Result := _TouchFileTimes(FileHandle, pSystemTime);
end;

type
  TImageDirectoryEntryToData = function (Base: Pointer; MappedAsImage: ByteBool;
    DirectoryEntry: USHORT; var Size: ULONG): Pointer; stdcall;

var
  _ImageDirectoryEntryToData: TImageDirectoryEntryToData = nil;

function ImageDirectoryEntryToData(Base: Pointer; MappedAsImage: ByteBool;
  DirectoryEntry: USHORT; var Size: ULONG): Pointer;
begin
  GetProcedureAddress(Pointer(@_ImageDirectoryEntryToData), ImageHlpLib, 'ImageDirectoryEntryToData');
  Result := _ImageDirectoryEntryToData(Base, MappedAsImage, DirectoryEntry, Size);
end;

type
  TImageRvaToSection = function (NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG): PImageSectionHeader; stdcall;

var
  _ImageRvaToSection: TImageRvaToSection = nil;

function ImageRvaToSection(NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG): PImageSectionHeader;
begin
  GetProcedureAddress(Pointer(@_ImageRvaToSection), ImageHlpLib, 'ImageRvaToSection');
  Result := _ImageRvaToSection(NtHeaders, Base, Rva);
end;

type
  TImageRvaToVa = function (NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG;
    LastRvaSection: PPImageSectionHeader): Pointer; stdcall;

var
  _ImageRvaToVa: TImageRvaToVa = nil;

function ImageRvaToVa(NtHeaders: PImageNtHeaders; Base: Pointer; Rva: ULONG;
  LastRvaSection: PPImageSectionHeader): Pointer;
begin
  GetProcedureAddress(Pointer(@_ImageRvaToVa), ImageHlpLib, 'ImageRvaToVa');
  Result := _ImageRvaToVa(NtHeaders, Base, Rva, LastRvaSection);
end;

const
  ntdll = 'ntdll.dll';

type
  TNtQueryInformationThread = function (ThreadHandle: THandle; ThreadInformationClass: THREAD_INFORMATION_CLASS;
    ThreadInformation: Pointer; ThreadInformationLength: ULONG; ReturnLength: PULONG): NTSTATUS; stdcall;
var
  _NtQueryInformationThread: TNtQueryInformationThread = nil;

function NtQueryInformationThread(ThreadHandle: THandle; ThreadInformationClass: THREAD_INFORMATION_CLASS;
  ThreadInformation: Pointer; ThreadInformationLength: ULONG; ReturnLength: PULONG): NTSTATUS;
begin
  GetProcedureAddress(Pointer(@_NtQueryInformationThread), ntdll, 'NtQueryInformationThread');
  Result := _NtQueryInformationThread(ThreadHandle, ThreadInformationClass, ThreadInformation, ThreadInformationLength, ReturnLength);
end;


type
  TCaptureStackBackTrace = function(FramesToSkip, FramesToCapture: DWORD;
    BackTrace: Pointer; out BackTraceHash: DWORD): Word; stdcall;

var
  _CaptureStackBackTrace: TCaptureStackBackTrace = nil;

function CaptureStackBackTrace(FramesToSkip, FramesToCapture: DWORD;
  BackTrace: Pointer; out BackTraceHash: DWORD): Word; stdcall;
begin
  GetProcedureAddress(Pointer(@_CaptureStackBackTrace), kernel32, 'RtlCaptureStackBackTrace');
  Result := _CaptureStackBackTrace(FramesToSkip, FramesToCapture, BackTrace, BackTraceHash);
end;


initialization  //10.05.2022 Profiling Startup Time (PAT) OK: 0 ms
  DebugInfoCritSect := TCriticalSection.Create;
  GlobalModulesList := TJclGlobalModulesList.Create;
  GlobalStackList := TJclGlobalStackList.Create;
  AddIgnoredException(EAbort);
  SetupExceptionProcs;

finalization
  ResetExceptionProcs;

  FreeAndNil(DebugInfoList);
  FreeAndNil(GlobalStackList);
  FreeAndNil(GlobalModulesList);
  FreeAndNil(DebugInfoCritSect);
  FreeAndNil(InfoSourceClassList);
  FreeAndNil(IgnoredExceptions);
  FreeAndNil(IgnoredExceptionClassNames);

  TJclDebugInfoSymbols.CleanupDebugSymbols;

end.
