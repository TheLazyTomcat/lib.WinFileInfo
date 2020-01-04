{-------------------------------------------------------------------------------

  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.

-------------------------------------------------------------------------------}
{===============================================================================

  WinFileInfo

    Main aim of this library is to provide a simple way of obtaining file
    information such as size, attributes, time of creation and, in case of
    binaries, a version information.
    A complete parsing of raw version information data is implemented, so
    it is possible to obtain information even from badly constructed version
    info resource.
    Although the library is writen only for Windows OS, it can be compiled for
    other systems. But in such case, it provides only a routine for conversion
    of file size into a textual representation (with proper units).

  Version 1.0.7 (2020-01-04)

  Last change 2020-01-04

  ©2015-2020 František Milt

  Contacts:
    František Milt: frantisek.milt@gmail.com

  Support:
    If you find this code useful, please consider supporting its author(s) by
    making a small donation using the following link(s):

      https://www.paypal.me/FMilt

  Changelog:
    For detailed changelog and history please refer to this git repository:

      github.com/TheLazyTomcat/Lib.WinFileInfo

  Dependencies:
    AuxTypes - github.com/TheLazyTomcat/Lib.AuxTypes
  * StrRect  - github.com/TheLazyTomcat/Lib.StrRect

    StrRect is currently required only when compiled for Windows OS.

===============================================================================}
unit WinFileInfo;

{$IF not(Defined(MSWINDOWS) or Defined(WINDOWS))}
  {$DEFINE LimitedImplementation}
{$ELSE}
  {$UNDEF LimitedImplementation}
{$IFEND}

{$IFDEF FPC}
  {
    Activate symbol BARE_FPC if you want to compile this unit outside of
    Lazarus.
    Non-unicode strings are assumed to be ANSI-encoded when defined, otherwise
    they are assumed to be UTF8-encoded.

    Not defined by default.
  }
  {.$DEFINE BARE_FPC}
  {$MODE ObjFPC}{$H+}
  {$DEFINE FPC_DisableWarns}
  {$MACRO ON}
{$ENDIF}

{$IF Defined(FPC) and not Defined(Unicode) and (FPC_FULLVERSION < 20701) and not Defined(BARE_FPC)}
  {$DEFINE UTF8Wrappers}
{$ELSE}
  {$UNDEF UTF8Wrappers}
{$IFEND}

interface

uses
  SysUtils
{$IFDEF LimitedImplementation}
  // non-win
  {$IFNDEF FPC}
  ,AuxTypes
  {$ENDIF}
{$ELSE}
  // windows
  , Windows, Classes,
  AuxTypes
{$ENDIF};

{===============================================================================
    Auxiliary functions
===============================================================================}

{
  FileSizeToStr expects passed number to be a file size and converts it to its
  string representation (as a decimal number if needed), including proper unit
  (KiB, MiB, ...).
}
Function FileSizeToStr(FileSize: UInt64; FormatSettings: TFormatSettings; SpaceUnit: Boolean = True): String; overload;
Function FileSizeToStr(FileSize: UInt64; SpaceUnit: Boolean = True): String; overload;

{$IFNDEF LimitedImplementation}
{===============================================================================
--------------------------------------------------------------------------------
                                  TWinFileInfo
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TWinFileInfo - library specific exception
===============================================================================}

type
  EWFIException = class(Exception);

  EWFISystemError      = class(EWFIException);
  EWFIIndexOutOfBounds = class(EWFIException);

{===============================================================================
    TWinFileInfo - system constants
===============================================================================}
const
  // File attributes flags
  INVALID_FILE_ATTRIBUTES = DWORD(-1); 

  FILE_ATTRIBUTE_ARCHIVE             = $20;
  FILE_ATTRIBUTE_COMPRESSED          = $800;
  FILE_ATTRIBUTE_DEVICE              = $40;
  FILE_ATTRIBUTE_DIRECTORY           = $10;
  FILE_ATTRIBUTE_ENCRYPTED           = $4000;
  FILE_ATTRIBUTE_HIDDEN              = $2;
  FILE_ATTRIBUTE_INTEGRITY_STREAM    = $8000;
  FILE_ATTRIBUTE_NORMAL              = $80;
  FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = $2000;
  FILE_ATTRIBUTE_NO_SCRUB_DATA       = $20000;
  FILE_ATTRIBUTE_OFFLINE             = $1000;
  FILE_ATTRIBUTE_READONLY            = $1;
  FILE_ATTRIBUTE_REPARSE_POINT       = $400;
  FILE_ATTRIBUTE_SPARSE_FILE         = $200;
  FILE_ATTRIBUTE_SYSTEM              = $4;
  FILE_ATTRIBUTE_TEMPORARY           = $100;
  FILE_ATTRIBUTE_VIRTUAL             = $10000;

  // Flags for field TVSFixedFileInfo.dwFileFlags
  VS_FF_DEBUG        = $00000001;
  VS_FF_INFOINFERRED = $00000010;
  VS_FF_PATCHED      = $00000004;
  VS_FF_PRERELEASE   = $00000002;
  VS_FF_PRIVATEBUILD = $00000008;
  VS_FF_SPECIALBUILD = $00000020;

  // Flags for field TVSFixedFileInfo.dwFileOS
  VOS_DOS           = $00010000;
  VOS_NT            = $00040000;
  VOS__WINDOWS16    = $00000001;
  VOS__WINDOWS32    = $00000004;
  VOS_OS216         = $00020000;
  VOS_OS232         = $00030000;
  VOS__PM16         = $00000002;
  VOS__PM32         = $00000003;
  VOS_UNKNOWN       = $00000000;
  VOS_DOS_WINDOWS16 = $00010001;
  VOS_DOS_WINDOWS32 = $00010004;
  VOS_NT_WINDOWS32  = $00040004;
  VOS_OS216_PM16    = $00020002;
  VOS_OS232_PM32    = $00030003;

  // Flags for field TVSFixedFileInfo.dwFileType
  VFT_APP        = $00000001;
  VFT_DLL        = $00000002;
  VFT_DRV        = $00000003;
  VFT_FONT       = $00000004;
  VFT_STATIC_LIB = $00000007;
  VFT_UNKNOWN    = $00000000;
  VFT_VXD        = $00000005;

  // Flags for field TVSFixedFileInfo.dwFileSubtype when
  // TVSFixedFileInfo.dwFileType is set to VFT_DRV
  VFT2_DRV_COMM              = $0000000A;
  VFT2_DRV_DISPLAY           = $00000004;
  VFT2_DRV_INSTALLABLE       = $00000008;
  VFT2_DRV_KEYBOARD          = $00000002;
  VFT2_DRV_LANGUAGE          = $00000003;
  VFT2_DRV_MOUSE             = $00000005;
  VFT2_DRV_NETWORK           = $00000006;
  VFT2_DRV_PRINTER           = $00000001;
  VFT2_DRV_SOUND             = $00000009;
  VFT2_DRV_SYSTEM            = $00000007;
  VFT2_DRV_VERSIONED_PRINTER = $0000000C;
  VFT2_UNKNOWN               = $00000000;

  // Flags for field TVSFixedFileInfo.dwFileSubtype when
  // TVSFixedFileInfo.dwFileType is set to VFT_FONT
  VFT2_FONT_RASTER   = $00000001;
  VFT2_FONT_TRUETYPE = $00000003;
  VFT2_FONT_VECTOR   = $00000002;

{===============================================================================
    TWinFileInfo - structures
===============================================================================}
{
  Following structures are used to store information about requested file in
  a more user-friendly and better accessible way.
}

type
  TWFIFileAttributesDecoded = record
    Archive:            Boolean;
    Compressed:         Boolean;
    Device:             Boolean;
    Directory:          Boolean;
    Encrypted:          Boolean;
    Hidden:             Boolean;
    IntegrityStream:    Boolean;
    Normal:             Boolean;
    NotContentIndexed:  Boolean;
    NoScrubData:        Boolean;
    Offline:            Boolean;
    ReadOnly:           Boolean;
    ReparsePoint:       Boolean;
    SparseFile:         Boolean;
    System:             Boolean;
    Temporary:          Boolean;
    Virtual:            Boolean;
  end;

//------------------------------------------------------------------------------
{
  Group of structures used to store decoded information from fixed file info
  part of version information resource.
}

  TWFIFixedFileInfo_VersionMembers = record
    Major:    UInt16;
    Minor:    UInt16;
    Release:  UInt16;
    Build:    UInt16;
  end;  

  TWFIFixedFileInfo_FileFlags = record
    Debug:        Boolean;
    InfoInferred: Boolean;
    Patched:      Boolean;
    Prerelease:   Boolean;
    PrivateBuild: Boolean;
    SpecialBuild: Boolean;
  end;

  TWFIFixedFileInfoDecoded = record
    FileVersionFull:        UInt64;
    FileVersionMembers:     TWFIFixedFileInfo_VersionMembers;
    FileVersionStr:         String;
    ProductVersionFull:     UInt64;
    ProductVersionMembers:  TWFIFixedFileInfo_VersionMembers;
    ProductVersionStr:      String;
    FileFlags:              TWFIFixedFileInfo_FileFlags;
    FileOSStr:              String;
    FileTypeStr:            String;
    FileSubTypeStr:         String;
    FileDateFull:           UInt64;
  end;

//------------------------------------------------------------------------------
{
  Following structures are used to hold partially parsed information from
  version information structure.
}
  TWFIVersionInfoStruct_String = record
    Address:    Pointer;
    Size:       TMemSize;
    Key:        WideString;
    ValueType:  Integer;
    ValueSize:  TMemSize;
    Value:      Pointer;
  end;

  TWFIVersionInfoStruct_StringTable = record
    Address:    Pointer;
    Size:       TMemSize;
    Key:        WideString;
    ValueType:  Integer;
    ValueSize:  TMemSize;
    Strings:    array of TWFIVersionInfoStruct_String;
  end;

  TWFIVersionInfoStruct_StringFileInfo = record
    Address:      Pointer;
    Size:         TMemSize;
    Key:          WideString;
    ValueType:    Integer;
    ValueSize:    TMemSize;
    StringTables: array of TWFIVersionInfoStruct_StringTable;
  end;

  TWFIVersionInfoStruct_Var = record
    Address:    Pointer;
    Size:       TMemSize;
    Key:        WideString;
    ValueType:  Integer;
    ValueSize:  TMemSize;
    Value:      Pointer;
  end;

  TWFIVersionInfoStruct_VarFileInfo = record
    Address:    Pointer;
    Size:       TMemSize;
    Key:        WideString;
    ValueType:  Integer;
    ValueSize:  TMemSize;
    Vars:       array of TWFIVersionInfoStruct_Var;
  end;

  TWFIVersionInfoStruct = record
    Address:            Pointer;
    Size:               TMemSize;
    Key:                WideString;
    ValueType:          Integer;
    ValueSize:          TMemSize;
    FixedFileInfo:      Pointer;
    FixedFileInfoSize:  TMemSize;
    StringFileInfos:    array of TWFIVersionInfoStruct_StringFileInfo;
    VarFileInfos:       array of TWFIVersionInfoStruct_VarFileInfo;
  end;

//------------------------------------------------------------------------------
{
  Following structures are used to store fully parsed information from version
  information structure.
}

  TWFITranslationItem = record
    LanguageName: String;
    LanguageStr:  String;
    case Integer of
      0: (Language:     UInt16;
          CodePage:     UInt16);
      1: (Translation:  UInt32);
  end;

  TWFIStringTableItem = record
    Key:    String;
    Value:  String;
  end;

  TWFIStringTable = record
    Translation:  TWFITranslationItem;
    Strings:      array of TWFIStringTableItem;
  end;

{===============================================================================
    TWinFileInfo - loading strategy
===============================================================================}
{
  Loading strategy determines what file information will be loaded and decoded
  or parsed.
  Only one operation cannot be affected by loading strategy and is always
  performed even when loading strategy indicates no operation - check whether
  the file actually exists.
}
type
  TWFILoadingStrategyAction = (
  {
    load size, times, attributes and other basic info
  }
    lsaLoadBasicInfo,
  {
    decode attributes and set size string

    requires lsaLoadBasicInfo
  }
    lsaDecodeBasicInfo,
  {
    load version info, also loads translations and strings
  }
    lsaLoadVersionInfo,
  {
    do low-level parsing of version info data and enumerates keys

    requires lsaLoadVersionInfo
  }
    lsaParseVersionInfo,
  {
    load fixed file info

    requires lsaLoadVersionInfo
  }
    lsaLoadFixedFileInfo,
  {
    decode fixed file info if present, has effect only if FFI is present
    (indicated by (f)VersionInfoFixedFileInfoPresent)

    requires lsaLoadFixedFileInfo
  }
    lsaDecodeFixedFileInfo,
  {
    when no key is successfully enumerated (see lsaParseVersionInfo),
    a predefined set of keys is used

    requires lsaParseVersionInfo
  }
    lsaVerInfoPredefinedKeys,
  {
    extract translations from parsed version info - might get some translation
    that normal translation loading (see lsaLoadVersionInfo) missed

    requires lsaParseVersionInfo
  }
    lsaVerInfoExtractTranslations);

  TWFILoadingStrategy = set of TWFILoadingStrategyAction;

// some predefined loading strategies (no need to define type of the set)
const
  WFI_LS_LoadNone = [];

  WFI_LS_BasicInfo = [lsaLoadBasicInfo,lsaDecodeBasicInfo];

  WFI_LS_FullInfo = WFI_LS_BasicInfo + [lsaLoadVersionInfo,lsaParseVersionInfo,
                    lsaLoadFixedFileInfo,lsaDecodeFixedFileInfo];

  WFI_LS_All = WFI_LS_FullInfo + [lsaVerInfoPredefinedKeys,lsaVerInfoExtractTranslations];

  WFI_LS_VersionInfo = [lsaLoadVersionInfo,lsaParseVersionInfo,lsaVerInfoExtractTranslations];

  WFI_LS_VersionInfoAndFFI = WFI_LS_VersionInfo + [lsaLoadFixedFileInfo,lsaDecodeFixedFileInfo];

{===============================================================================
    TWinFileInfo - class declaration
===============================================================================}
type
  TWinFileInfo = class(TObject)
  private
    // internals
    fLoadingStrategy:         TWFILoadingStrategy;
    fFormatSettings:          TFormatSettings;
    // basic initial file info
    fLongName:                String;
    fShortName:               String;
    fExists:                  Boolean;
    fFileHandle:              THandle;
    // basic loaded file info
    fSize:                    UInt64;
    fSizeStr:                 String;
    fCreationTime:            TDateTime;
    fLastAccessTime:          TDateTime;
    fLastWriteTime:           TDateTime;
    fNumberOfLinks:           UInt32;
  {
    combination of volume serial and file id can be used to determine file
    path equality (also for directories, but WFI does not support dirs)
  }
    fVolumeSerialNumber:      UInt32;
    fFileID:                  UInt64;
    // file attributes (part of basic info)
    fAttributesFlags:         DWORD;
    fAttributesStr:           String;
    fAttributesText:          String;
    fAttributesDecoded:       TWFIFileAttributesDecoded;
    // version info unparsed data
    fVerInfoSize:             TMemSize;
    fVerInfoData:             Pointer;
    // version info data
    fVersionInfoPresent:      Boolean;
    // version info - fixed file info
    fVersionInfoFFIPresent:   Boolean;
    fVersionInfoFFI:          TVSFixedFileInfo;
    fVersionInfoFFIDecoded:   TWFIFixedFileInfoDecoded;
    // version info partially parsed data
    fVersionInfoStruct:       TWFIVersionInfoStruct;
    // version info fully parsed data
    fVersionInfoParsed:       Boolean;
    fVersionInfoStringTables: array of TWFIStringTable;
    // getters for fVersionInfoStruct fields
    Function GetVersionInfoStringTableCount: Integer;
    Function GetVersionInfoStringTable(Index: Integer): TWFIStringTable;
    Function GetVersionInfoTranslationCount: Integer;
    Function GetVersionInfoTranslation(Index: Integer): TWFITranslationItem;
    Function GetVersionInfoStringCount(Table: Integer): Integer;
    Function GetVersionInfoString(Table,Index: Integer): TWFIStringTableItem;
    Function GetVersionInfoValue(const Language,Key: String): String;
  protected
    // version info loading methods
    procedure VersionInfo_LoadTranslations; virtual;
    procedure VersionInfo_LoadStrings; virtual;
    // version info parsing methods
    procedure VersionInfo_Parse; virtual;
    procedure VersionInfo_ExtractTranslations; virtual;    
    procedure VersionInfo_EnumerateKeys; virtual;
    // loading and decoding methods
    procedure LoadBasicInfo; virtual;
    procedure DecodeBasicInfo; virtual;
    procedure LoadVersionInfo; virtual;
    procedure LoadFixedFileInfo; virtual;
    procedure DecodeFixedFileInfo; virtual;
    // other protected methods
    procedure Clear; virtual;    
    procedure Initialize(const FileName: String); virtual;
    procedure Finalize; virtual;
  public
    constructor Create(LoadingStrategy: TWFILoadingStrategy = WFI_LS_All); overload;
    constructor Create(const FileName: String; LoadingStrategy: TWFILoadingStrategy = WFI_LS_All); overload;
    destructor Destroy; override;
    procedure Refresh; virtual;
    Function IndexOfVersionInfoStringTable(Translation: DWORD): Integer; virtual;
    Function IndexOfVersionInfoString(Table: Integer; const Key: String): Integer; virtual;
    procedure CreateReport(Strings: TStrings); overload; virtual;
    Function CreateReport: String; overload; virtual;
    // internals
    property LoadingStrategy: TWFILoadingStrategy read fLoadingStrategy write fLoadingStrategy;
    property FormatSettings: TFormatSettings read fFormatSettings write fFormatSettings;
    // basic initial file info
    property Name: String read fLongName;
    property LongName: String read fLongName;
    property ShortName: String read fShortName;
    property Exists: Boolean read fExists;
    property FileHandle: THandle read fFileHandle;
    // basic loaded file info
    property Size: UInt64 read fSize;
    property SizeStr: String read fSizeStr;
    property CreationTime: TDateTime read fCreationTime;
    property LastAccessTime: TDateTime read fLastAccessTime;
    property LastWriteTime: TDateTime read fLastWriteTime;
    property NumberOfLinks: UInt32 read fNumberOfLinks;
    property VolumeSerialNumber: UInt32 read fVolumeSerialNumber;
    property FileID: UInt64 read fFileID;
    // file attributes (part of basic info)
    property AttributesFlags: DWORD read fAttributesFlags;
    property AttributesStr: String read fAttributesStr;
    property AttributesText: String read fAttributesText;
    property AttributesDecoded: TWFIFileAttributesDecoded read fAttributesDecoded;
    // version info unparsed data
    property VerInfoSize: PtrUInt read fVerInfoSize;
    property VerInfoData: Pointer read fVerInfoData;
    // version info data
    property VersionInfoPresent: Boolean read fVersionInfoPresent;
    // version info - fixed file info 
    property VersionInfoFixedFileInfoPresent: Boolean read fVersionInfoFFIPresent;
    property VersionInfoFixedFileInfo: TVSFixedFileInfo read fVersionInfoFFI;
    property VersionInfoFixedFileInfoDecoded: TWFIFixedFileInfoDecoded read fVersionInfoFFIDecoded;
    // version info partially parsed data
    property VersionInfoStruct: TWFIVersionInfoStruct read fVersionInfoStruct;
    // version info fully parsed data
    property VersionInfoParsed: Boolean read fVersionInfoParsed;
    property VersionInfoStringTableCount: Integer read GetVersionInfoStringTableCount;
    property VersionInfoStringTables[Index: Integer]: TWFIStringTable read GetVersionInfoStringTable;
    property VersionInfoTranslationCount: Integer read GetVersionInfoTranslationCount;
    property VersionInfoTranslations[Index: Integer]: TWFITranslationItem read GetVersionInfoTranslation;
    property VersionInfoStringCount[Table: Integer]: Integer read GetVersionInfoStringCount;
    property VersionInfoStrings[Table,Index: Integer]: TWFIStringTableItem read GetVersionInfoString;
    property VersionInfoValues[const Language,Key: String]: String read GetVersionInfoValue; default;
  end;

{$ENDIF LimitedImplementation}

implementation

{$IFNDEF LimitedImplementation}
uses
  {$IFDEF UTF8Wrappers}LazFileUtils,{$ENDIF}
  StrRect;
{$ENDIF LimitedImplementation}

{$IFDEF FPC_DisableWarns}
  {$DEFINE FPCDWM}
  {$DEFINE W4055:={$WARN 4055 OFF}}   // Conversion between ordinals and pointers is not portable
  {$DEFINE W5024:={$WARN 5024 OFF}}   // Parameter "$1" not used
  {$DEFINE W5057:={$WARN 5057 OFF}}   // Local variable "$1" does not seem to be initialized
  {$PUSH}{$WARN 2005 OFF}             // Comment level $1 found
  {$IF Defined(FPC) and (FPC_FULLVERSION >= 30000)}
    {$DEFINE W5091:={$WARN 5091 OFF}} // Local variable "$1" of a managed type does not seem to be initialized
  {$ELSE}
    {$DEFINE W5091:=}
  {$IFEND}
  {$POP}
{$ENDIF}

{===============================================================================
    Auxiliary functions
===============================================================================}
{-------------------------------------------------------------------------------
    Auxiliary functions - public functions
-------------------------------------------------------------------------------}

Function FileSizeToStr(FileSize: UInt64; FormatSettings: TFormatSettings; SpaceUnit: Boolean = True): String;
const
  BinaryPrefix: array[0..8] of String = ('','Ki','Mi','Gi','Ti','Pi','Ei','Zi','Yi');
  PrefixShift = 10;
var
  Offset: Integer;  
  Deci:   Integer;  // number of shown decimal places
  Num:    Double;
begin
Offset := -1;
repeat
  Inc(Offset);
until ((FileSize shr (PrefixShift * Succ(Offset))) = 0) or (Offset >= 8);
case FileSize shr (PrefixShift * Offset) of
   1..9:  Deci := 2;
  10..99: Deci := 1;
else
  Deci := 0;
end;
Num := (FileSize shr (PrefixShift * Offset));
If Offset > 0 then
  Num := Num + (((FileSize shr (PrefixShift * Pred(Offset))) and 1023) / 1024)
else
  Deci := 0;
If SpaceUnit then
  Result := Format('%.*f %sB',[Deci,Num,BinaryPrefix[Offset]],FormatSettings)
else
  Result := Format('%.*f%sB',[Deci,Num,BinaryPrefix[Offset]],FormatSettings)
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

{$IFDEF FPCDWM}{$PUSH}W5057 W5091{$ENDIF}
Function FileSizeToStr(FileSize: UInt64; SpaceUnit: Boolean = True): String;
var
  FormatSettings: TFormatSettings;
begin
{$WARN SYMBOL_PLATFORM OFF}
{$IF not Defined(FPC) and (CompilerVersion >= 18)} // Delphi 2006+
FormatSettings := TFormatSettings.Create(LOCALE_USER_DEFAULT);
{$ELSE}
{$IFDEF LimitedImplementation}
// non-windows
FormatSettings := DefaultFormatSettings;
{$ELSE}
// windows
GetLocaleFormatSettings(LOCALE_USER_DEFAULT,FormatSettings);
{$ENDIF}
{$IFEND}
{$WARN SYMBOL_PLATFORM ON}
Result := FileSizeToStr(FileSize,FormatSettings,SpaceUnit);
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

{$IFNDEF LimitedImplementation}
{-------------------------------------------------------------------------------
    Auxiliary functions - local functions
-------------------------------------------------------------------------------}

{$IF not Declared(CP_THREAD_ACP)}
const
  CP_THREAD_ACP = 3;
{$IFEND}

{$IFDEF FPCDWM}{$PUSH}W5024{$ENDIF}
Function WideToString(const WStr: WideString; AnsiCodePage: UINT = CP_THREAD_ACP): String;
begin
{$IFDEF Unicode}
// unicode Delphi or FPC (String = UnicodeString)
Result := WStr;
{$ELSE}
// non-unicode...
{$IF Defined(FPC) and not Defined(BARE_FPC)}
// FPC in Lazarus (String = UTF8String)
Result := UTF8Encode(WStr)
{$ELSE}
// bare FPC or Delphi (String = AnsiString)
SetLength(Result,WideCharToMultiByte(AnsiCodePage,0,PWideChar(WStr),Length(WStr),nil,0,nil,nil));
WideCharToMultiByte(AnsiCodePage,0,PWideChar(WStr),Length(WStr),PAnsiChar(Result),Length(Result) * SizeOf(AnsiChar),nil,nil);
// A wrong codepage might be stored, try translation with default cp
If (AnsiCodePage <> CP_THREAD_ACP) and (Length(Result) <= 0) and (Length(WStr) > 0) then
  Result := WideToString(WStr);
{$IFEND}
{$ENDIF}
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

{===============================================================================
--------------------------------------------------------------------------------
                                  TWinFileInfo
--------------------------------------------------------------------------------
===============================================================================}
{===============================================================================
    TWinFileInfo - external functions
===============================================================================}

{$IF not Declared(GetFileSizeEx)}
Function GetFileSizeEx(hFile: THandle; lpFileSize: PUInt64): BOOL; stdcall; external 'kernel32.dll';
{$IFEND}

{===============================================================================
    TWinFileInfo - conversion tables
===============================================================================}

// structures used in conversion tables as items
type
  TWFIAttributeString = record
    Flag: DWORD;
    Text: String;
    Str:  String;
  end;

  TWFIFlagText = record
    Flag: DWORD;
    Text: String;
  end;

//------------------------------------------------------------------------------
{
  tables used to convert some binary information (mainly flags) to a textual
  representation
}
const
  WFI_FILE_ATTR_STRS: array[0..16] of TWFIAttributeString = (
    (Flag: FILE_ATTRIBUTE_ARCHIVE;             Text: 'Archive';             Str: 'A'),
    (Flag: FILE_ATTRIBUTE_COMPRESSED;          Text: 'Compressed';          Str: 'C'),
    (Flag: FILE_ATTRIBUTE_DEVICE;              Text: 'Device';              Str: ''),
    (Flag: FILE_ATTRIBUTE_DIRECTORY;           Text: 'Directory';           Str: 'D'),
    (Flag: FILE_ATTRIBUTE_ENCRYPTED;           Text: 'Encrypted';           Str: 'E'),
    (Flag: FILE_ATTRIBUTE_HIDDEN;              Text: 'Hidden';              Str: 'H'),
    (Flag: FILE_ATTRIBUTE_INTEGRITY_STREAM;    Text: 'Integrity stream';    Str: ''),
    (Flag: FILE_ATTRIBUTE_NORMAL;              Text: 'Normal';              Str: 'N'),
    (Flag: FILE_ATTRIBUTE_NOT_CONTENT_INDEXED; Text: 'Not content indexed'; Str: 'I'),
    (Flag: FILE_ATTRIBUTE_NO_SCRUB_DATA;       Text: 'No scrub data';       Str: ''),
    (Flag: FILE_ATTRIBUTE_OFFLINE;             Text: 'Offline';             Str: 'O'),
    (Flag: FILE_ATTRIBUTE_READONLY;            Text: 'Read only';           Str: 'R'),
    (Flag: FILE_ATTRIBUTE_REPARSE_POINT;       Text: 'Reparse point';       Str: 'L'),
    (Flag: FILE_ATTRIBUTE_SPARSE_FILE;         Text: 'Sparse file';         Str: 'P'),
    (Flag: FILE_ATTRIBUTE_SYSTEM;              Text: 'System';              Str: 'S'),
    (Flag: FILE_ATTRIBUTE_TEMPORARY;           Text: 'Temporary';           Str: 'T'),
    (Flag: FILE_ATTRIBUTE_VIRTUAL;             Text: 'Virtual';             Str: ''));

//------------------------------------------------------------------------------

  WFI_FFI_FILE_OS_STRS: array[0..13] of TWFIFlagText = (
    (Flag: VOS_DOS;           Text: 'MS-DOS'),
    (Flag: VOS_NT;            Text: 'Windows NT'),
    (Flag: VOS__WINDOWS16;    Text: '16-bit Windows'),
    (Flag: VOS__WINDOWS32;    Text: '32-bit Windows'),
    (Flag: VOS_OS216;         Text: '16-bit OS/2'),
    (Flag: VOS_OS232;         Text: '32-bit OS/2'),
    (Flag: VOS__PM16;         Text: '16-bit Presentation Manager'),
    (Flag: VOS__PM32;         Text: '32-bit Presentation Manager'),
    (Flag: VOS_UNKNOWN;       Text: 'Unknown'),
    (Flag: VOS_DOS_WINDOWS16; Text: '16-bit Windows running on MS-DOS'),
    (Flag: VOS_DOS_WINDOWS32; Text: '32-bit Windows running on MS-DOS'),
    (Flag: VOS_NT_WINDOWS32;  Text: 'Windows NT'),
    (Flag: VOS_OS216_PM16;    Text: '16-bit Presentation Manager running on 16-bit OS/2'),
    (Flag: VOS_OS232_PM32;    Text: '32-bit Presentation Manager running on 32-bit OS/2'));

//------------------------------------------------------------------------------

  WFI_FFI_FILE_TYPE_STRS: array[0..6] of TWFIFlagText = (
    (Flag: VFT_APP;        Text: 'Application'),
    (Flag: VFT_DLL;        Text: 'DLL'),
    (Flag: VFT_DRV;        Text: 'Device driver'),
    (Flag: VFT_FONT;       Text: 'Font'),
    (Flag: VFT_STATIC_LIB; Text: 'Static-link library'),
    (Flag: VFT_UNKNOWN;    Text: 'Unknown'),
    (Flag: VFT_VXD;        Text: 'Virtual device'));

//------------------------------------------------------------------------------

  WFI_FFI_FILE_SUBTYPE_DRV_STRS: array[0..11] of TWFIFlagText = (
    (Flag: VFT2_DRV_COMM;              Text: 'Communications driver'),
    (Flag: VFT2_DRV_DISPLAY;           Text: 'Display driver'),
    (Flag: VFT2_DRV_INSTALLABLE;       Text: 'Installable driver'),
    (Flag: VFT2_DRV_KEYBOARD;          Text: 'Keyboard driver'),
    (Flag: VFT2_DRV_LANGUAGE;          Text: 'Language driver'),
    (Flag: VFT2_DRV_MOUSE;             Text: 'Mouse driver'),
    (Flag: VFT2_DRV_NETWORK;           Text: 'Network driver'),
    (Flag: VFT2_DRV_PRINTER;           Text: 'Printer driver'),
    (Flag: VFT2_DRV_SOUND;             Text: 'Sound driver'),
    (Flag: VFT2_DRV_SYSTEM;            Text: 'System driver'),
    (Flag: VFT2_DRV_VERSIONED_PRINTER; Text: 'Versioned printer driver'),
    (Flag: VFT2_UNKNOWN;               Text: 'Unknown'));

//------------------------------------------------------------------------------

  WFI_FFI_FILE_SUBTYPE_FONT_STRS: array[0..3] of TWFIFlagText = (
    (Flag: VFT2_FONT_RASTER;   Text: 'Raster font'),
    (Flag: VFT2_FONT_TRUETYPE; Text: 'TrueType font'),
    (Flag: VFT2_FONT_VECTOR;   Text: 'Vector font'),
    (Flag: VFT2_UNKNOWN;       Text: 'Unknown'));

//------------------------------------------------------------------------------

  WFI_VERINFO_PREDEF_KEYS: array[0..11] of String = (
    'Comments','CompanyName','FileDescription','FileVersion','InternalName',
    'LegalCopyright','LegalTrademarks','OriginalFilename','ProductName',
    'ProductVersion','PrivateBuild','SpecialBuild');

{===============================================================================
    TWinFileInfo - class implementation
===============================================================================}
{-------------------------------------------------------------------------------
    TWinFileInfo - private methods
-------------------------------------------------------------------------------}

Function TWinFileInfo.GetVersionInfoStringTableCount: Integer;
begin
Result := Length(fVersionInfoStringTables);
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.GetVersionInfoStringTable(Index: Integer): TWFIStringTable;
begin
If (Index >= Low(fVersionInfoStringTables)) and (Index <= High(fVersionInfoStringTables)) then
  Result := fVersionInfoStringTables[Index]
else
  raise EWFIIndexOutOfBounds.CreateFmt('TWinFileInfo.GetVersionInfoStringTable: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.GetVersionInfoTranslationCount: Integer;
begin
Result := Length(fVersionInfoStringTables);
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.GetVersionInfoTranslation(Index: Integer): TWFITranslationItem;
begin
If (Index >= Low(fVersionInfoStringTables)) and (Index <= High(fVersionInfoStringTables)) then
  Result := fVersionInfoStringTables[Index].Translation
else
  raise EWFIIndexOutOfBounds.CreateFmt('TWinFileInfo.GetVersionInfoTranslation: Index (%d) out of bounds.',[Index]);
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.GetVersionInfoStringCount(Table: Integer): Integer;
begin
Result := Length(GetVersionInfoStringTable(Table).Strings);
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.GetVersionInfoString(Table,Index: Integer): TWFIStringTableItem;
begin
with GetVersionInfoStringTable(Table) do
  begin
    If (Index >= Low(Strings)) and (Index <= High(Strings)) then
      Result := Strings[Index]
    else
      raise EWFIIndexOutOfBounds.CreateFmt('TWinFileInfo.GetVersionInfoString: Index (%d) out of bounds.',[Index]);
  end;
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
Function TWinFileInfo.GetVersionInfoValue(const Language,Key: String): String;
var
  StrPtr:   Pointer;
  StrSize:  UInt32;
begin
Result := '';
If fVersionInfoPresent and (Language <> '') and (Key <> '') then
  If VerQueryValue(fVerInfoData,PChar(Format('\StringFileInfo\%s\%s',[Language,Key])),StrPtr,StrSize) then
    Result := WinToStr(PChar(StrPtr));
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

{-------------------------------------------------------------------------------
    TWinFileInfo - protected methods
-------------------------------------------------------------------------------}

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
procedure TWinFileInfo.VersionInfo_LoadTranslations;
var
  TrsPtr:   Pointer;
  TrsSize:  UInt32;
  i:        Integer;
begin
If VerQueryValue(fVerInfoData,'\VarFileInfo\Translation',TrsPtr,TrsSize) then
  begin
    SetLength(fVersionInfoStringTables,TrsSize div SizeOf(UInt32));
    For i := Low(fVersionInfoStringTables) to High(fVersionInfoStringTables) do
      with fVersionInfoStringTables[i].Translation do
        begin
        {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
          Translation := PUInt32(PtrUInt(TrsPtr) + (PtrUInt(i) * SizeOf(UInt32)))^;
        {$IFDEF FPCDWM}{$POP}{$ENDIF}
          SetLength(LanguageName,256);  // should be sufficiently long enough, hopefully
          SetLength(LanguageName,VerLanguageName(Translation,PChar(LanguageName),Length(LanguageName)));
          LanguageName := WinToStr(LanguageName);
          LanguageStr := AnsiUpperCase(IntToHex(Language,4) + IntToHex(CodePage,4));
        end;
  end;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
procedure TWinFileInfo.VersionInfo_LoadStrings;
var
  Table:    Integer;
  i,j:      Integer;
  StrPtr:   Pointer;
  StrSize:  UInt32;
begin
// loads value of string according to the key
For Table := Low(fVersionInfoStringTables) to High(fVersionInfoStringTables) do
  with fVersionInfoStringTables[Table] do
    For i := High(Strings) downto Low(Strings) do
      If not VerQueryValue(fVerInfoData,PChar(Format('\StringFileInfo\%s\%s',[Translation.LanguageStr,StrToWin(Strings[i].Key)])),StrPtr,StrSize) then
        begin
          // remove this one string as its value could not be obtained
          For j := i to Pred(High(Strings)) do
            Strings[j] := Strings[j + 1];
          SetLength(Strings,Length(Strings) - 1);
        end
      else Strings[i].Value := WinToStr(PChar(StrPtr))
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

procedure TWinFileInfo.VersionInfo_Parse;
type
  PVIS_Base = ^TVIS_Base;
  TVIS_Base = record
    Address:    Pointer;
    Size:       TMemSize;
    Key:        WideString;
    ValueType:  Integer;
    ValueSize:  TMemSize;
  end;
var
  CurrentAddress: Pointer;
  TempBlock:      TVIS_Base;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

  Function Align32bit(Ptr: Pointer): Pointer;
  begin
  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
    If ((PtrUInt(Ptr) and 3) <> 0) then
      Result := Pointer((PtrUInt(Ptr) and not PtrUInt(3)) + 4)
    else
      Result := Ptr;
  {$IFDEF FPCDWM}{$POP}{$ENDIF}
  end;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

  procedure ParseBlock(var Ptr: Pointer; BlockBase: Pointer);
  begin
    PVIS_Base(BlockBase)^.Address := Ptr;
    PVIS_Base(BlockBase)^.Size := PUInt16(PVIS_Base(BlockBase)^.Address)^;
  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
    PVIS_Base(BlockBase)^.Key := PWideChar(PtrUInt(PVIS_Base(BlockBase)^.Address) + (3 * SizeOf(UInt16)));
    PVIS_Base(BlockBase)^.ValueType := PUInt16(PtrUInt(PVIS_Base(BlockBase)^.Address) + (2 * SizeOf(UInt16)))^;
    PVIS_Base(BlockBase)^.ValueSize := PUInt16(PtrUInt(PVIS_Base(BlockBase)^.Address) + SizeOf(UInt16))^;
    Ptr := Align32bit(Pointer(PtrUInt(PVIS_Base(BlockBase)^.Address) + (3 * SizeOf(UInt16)) +
             PtrUInt((Length(PVIS_Base(BlockBase)^.Key) + 1{terminating zero}) * SizeOf(WideChar))));
  {$IFDEF FPCDWM}{$POP}{$ENDIF}
  end;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

  Function CheckPointer(var Ptr: Pointer; BlockBase: Pointer): Boolean;
  begin
  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
    Result := (PtrUInt(Ptr) >= PtrUInt(PVIS_Base(BlockBase)^.Address)) and
              (PtrUInt(Ptr) < (PtrUInt(PVIS_Base(BlockBase)^.Address) + PVIS_Base(BlockBase)^.Size));
  {$IFDEF FPCDWM}{$POP}{$ENDIF}
  end;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

begin
If (fVerInfoSize >= 6) and (fVerInfoSize >= PUInt16(fVerInfoData)^) then
  try
    CurrentAddress := fVerInfoData;
    ParseBlock(CurrentAddress,@fVersionInfoStruct);
  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
    fVersionInfoStruct.FixedFileInfoSize := PUInt16(PtrUInt(fVersionInfoStruct.Address) + SizeOf(UInt16))^;
    fVersionInfoStruct.FixedFileInfo := CurrentAddress;
    CurrentAddress := Align32bit(Pointer(PtrUInt(CurrentAddress) + fVersionInfoStruct.FixedFileInfoSize));
  {$IFDEF FPCDWM}{$POP}{$ENDIF}
    // traverse remaining memory as long as the the current pointer points to a valid memory area
    while CheckPointer(CurrentAddress,@fVersionInfoStruct) do
      begin
        ParseBlock(CurrentAddress,@TempBlock);
        If WideSameText(TempBlock.Key,WideString('StringFileInfo')) then
          begin
            // strings
            SetLength(fVersionInfoStruct.StringFileInfos,Length(fVersionInfoStruct.StringFileInfos) + 1);
            with fVersionInfoStruct.StringFileInfos[High(fVersionInfoStruct.StringFileInfos)] do
              begin
                Address := TempBlock.Address;
                Size := TempBlock.Size;
                Key := TempBlock.Key;
                ValueType := TempBlock.ValueType;
                ValueSize := TempBlock.ValueSize;
                // parse-out string tables
                while CheckPointer(CurrentAddress,@fVersionInfoStruct.StringFileInfos[High(fVersionInfoStruct.StringFileInfos)]) do
                  begin
                    SetLength(StringTables,Length(StringTables) + 1);
                    ParseBlock(CurrentAddress,@StringTables[High(StringTables)]);
                    // parse-out individual strings
                    while CheckPointer(CurrentAddress,@StringTables[High(StringTables)]) do
                      with StringTables[High(StringTables)] do
                        begin
                          SetLength(Strings,Length(Strings) + 1);
                          ParseBlock(CurrentAddress,@Strings[High(Strings)]);
                          Strings[High(Strings)].Value := CurrentAddress;
                        {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
                          CurrentAddress := Align32bit(Pointer(PtrUInt(Strings[High(Strings)].Address) + Strings[High(Strings)].Size));
                        {$IFDEF FPCDWM}{$POP}{$ENDIF}
                        end;
                  end;
              end
          end
        else If WideSameText(TempBlock.Key,WideString('VarFileInfo')) then
          begin
            // variables
            SetLength(fVersionInfoStruct.VarFileInfos,Length(fVersionInfoStruct.VarFileInfos) + 1);
            with fVersionInfoStruct.VarFileInfos[High(fVersionInfoStruct.VarFileInfos)] do
              begin
                Address := TempBlock.Address;
                Size := TempBlock.Size;
                Key := TempBlock.Key;
                ValueType := TempBlock.ValueType;
                ValueSize := TempBlock.ValueSize;
                // parse-out variables
                while CheckPointer(CurrentAddress,@fVersionInfoStruct.VarFileInfos[High(fVersionInfoStruct.VarFileInfos)]) do
                  begin
                    SetLength(Vars,Length(Vars) + 1);
                    ParseBlock(CurrentAddress,@Vars[High(Vars)]);
                    Vars[High(Vars)].Value := CurrentAddress;
                  {$IFDEF FPCDWM}{$PUSH}W4055{$ENDIF}
                    CurrentAddress := Align32bit(Pointer(PtrUInt(Vars[High(Vars)].Address) + Vars[High(Vars)].Size));
                  {$IFDEF FPCDWM}{$POP}{$ENDIF}
                  end;
              end;
          end
        else raise EWFIException.CreateFmt('TWinFileInfo.VersionInfo_Parse: Unknown block (%s).',[TempBlock.Key]);
      end;
    fVersionInfoParsed := True;
  except
    fVersionInfoParsed := False;
    fVersionInfoStruct.Key := '';
    SetLength(fVersionInfoStruct.StringFileInfos,0);
    SetLength(fVersionInfoStruct.VarFileInfos,0);
    FillChar(fVersionInfoStruct,SizeOf(fVersionInfoStruct),0);
  end
else fVersionInfoParsed := False;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.VersionInfo_ExtractTranslations;
var
  i,Table:  Integer;

  Function TranslationIsListed(const LanguageStr: String): Boolean;
  var
    ii: Integer;
  begin
    Result := False;
    For ii := Low(fVersionInfoStringTables) to High(fVersionInfoStringTables) do
      If AnsiSameText(fVersionInfoStringTables[ii].Translation.LanguageStr,LanguageStr) then
        begin
          Result := True;
          Break;
        end;
  end;

begin
For i := Low(fVersionInfoStruct.StringFileInfos) to High(fVersionInfoStruct.StringFileInfos) do
  If WideSameText(fVersionInfoStruct.StringFileInfos[i].Key,WideString('StringFileInfo')) then
    For Table := Low(fVersionInfoStruct.StringFileInfos[i].StringTables) to High(fVersionInfoStruct.StringFileInfos[i].StringTables) do
      If not TranslationIsListed(WideToStr(fVersionInfoStruct.StringFileInfos[i].StringTables[Table].Key)) then
        begin
          SetLength(fVersionInfoStringTables,Length(fVersionInfoStringTables) + 1);
          with fVersionInfoStringTables[High(fVersionInfoStringTables)].Translation do
            begin
              LanguageStr := WideToStr(fVersionInfoStruct.StringFileInfos[i].StringTables[Table].Key);
              Language := StrToIntDef('$' + Copy(LanguageStr,1,4),0);
              CodePage := StrToIntDef('$' + Copy(LanguageStr,5,4),0);
              SetLength(LanguageName,256);
              SetLength(LanguageName,VerLanguageName(Translation,PChar(LanguageName),Length(LanguageName)));
              LanguageName := WinToStr(LanguageName);
            end;
        end;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.VersionInfo_EnumerateKeys;
var
  Table:  Integer;
  i,j,k:  Integer;
begin
For i := Low(fVersionInfoStruct.StringFileInfos) to High(fVersionInfoStruct.StringFileInfos) do
  If WideSameText(fVersionInfoStruct.StringFileInfos[i].Key,WideString('StringFileInfo')) then
    For Table := Low(fVersionInfoStringTables) to High(fVersionInfoStringTables) do
      with fVersionInfoStruct.StringFileInfos[i] do
        begin
          For j := Low(StringTables) to High(StringTables) do
            If WideSameText(StringTables[j].Key,StrToWide(fVersionInfoStringTables[Table].Translation.LanguageStr)) then
              begin
                SetLength(fVersionInfoStringTables[Table].Strings,Length(StringTables[j].Strings));
                For k := Low(StringTables[j].Strings) to High(StringTables[j].Strings) do
                  fVersionInfoStringTables[Table].Strings[k].Key := WideToString(StringTables[j].Strings[k].Key,fVersionInfoStringTables[Table].Translation.CodePage);
              end;
          If (Length(fVersionInfoStringTables[Table].Strings) <= 0) and (lsaVerInfoPredefinedKeys in fLoadingStrategy) then
            begin
              SetLength(fVersionInfoStringTables[Table].Strings,Length(WFI_VERINFO_PREDEF_KEYS));
              For j := Low(WFI_VERINFO_PREDEF_KEYS) to High(WFI_VERINFO_PREDEF_KEYS) do
                fVersionInfoStringTables[Table].Strings[j].Key := WFI_VERINFO_PREDEF_KEYS[j];
            end;
        end;
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
procedure TWinFileInfo.LoadBasicInfo;
var
  Info: TByHandleFileInformation;

  Function FileTimeToDateTime(FileTime: TFileTime): TDateTime;
  var
    LocalTime:  TFileTime;
    SystemTime: TSystemTime;
  begin
    If FileTimeToLocalFileTime(FileTime,LocalTime) then
      begin
        If FileTimeToSystemTime(LocalTime,SystemTime) then
          Result := SystemTimeToDateTime(SystemTime)
        else raise EWFISystemError.CreateFmt('FileTimeToSystemTime failed with error 0x%.8x.',[GetLastError]);
      end
    else raise EWFISystemError.CreateFmt('FileTimeToLocalFileTime failed with error 0x%.8x.',[GetLastError]);
  end;

begin
If GetFileInformationByHandle(fFileHandle,Info) then
  begin
    fSize := (UInt64(Info.nFileSizeHigh) shl 32) or UInt64(Info.nFileSizeLow);
    fCreationTime := FileTimeToDateTime(Info.ftCreationTime);
    fLastAccessTime := FileTimeToDateTime(Info.ftLastAccessTime);
    fLastWriteTime := FileTimeToDateTime(Info.ftLastWriteTime);
    fNumberOfLinks := Info.nNumberOfLinks;
    fVolumeSerialNumber := Info.dwVolumeSerialNumber;
    fFileID := (UInt64(Info.nFileIndexHigh) shl 32) or UInt64(Info.nFileIndexLow);
    fAttributesFlags := Info.dwFileAttributes;
  end
else raise EWFISystemError.CreateFmt('GetFileInformationByHandle failed with error 0x%.8x.',[GetLastError]);
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

procedure TWinFileInfo.DecodeBasicInfo;

  // This function also fills AttributesStr and AttributesText strings
  Function ProcessAttribute(AttributeFlag: DWORD): Boolean;
  var
    i:  Integer;
  begin
    Result := (fAttributesFlags and AttributeFlag) <> 0;
    If Result then
      For i := Low(WFI_FILE_ATTR_STRS) to High(WFI_FILE_ATTR_STRS) do
        If WFI_FILE_ATTR_STRS[i].Flag = AttributeFlag then
          begin
            fAttributesStr := fAttributesStr + WFI_FILE_ATTR_STRS[i].Str;
            If fAttributesText = '' then
              fAttributesText := WFI_FILE_ATTR_STRS[i].Text
            else
              fAttributesText := Format('%s, %s',[fAttributesText,WFI_FILE_ATTR_STRS[i].Text]);
            Break;
          end;
  end;

begin
fSizeStr := FileSizeToStr(fSize,fFormatSettings);
// attributes
fAttributesDecoded.Archive           := ProcessAttribute(FILE_ATTRIBUTE_ARCHIVE);
fAttributesDecoded.Compressed        := ProcessAttribute(FILE_ATTRIBUTE_COMPRESSED);
fAttributesDecoded.Device            := ProcessAttribute(FILE_ATTRIBUTE_DEVICE);
fAttributesDecoded.Directory         := ProcessAttribute(FILE_ATTRIBUTE_DIRECTORY);
fAttributesDecoded.Encrypted         := ProcessAttribute(FILE_ATTRIBUTE_ENCRYPTED);
fAttributesDecoded.Hidden            := ProcessAttribute(FILE_ATTRIBUTE_HIDDEN);
fAttributesDecoded.IntegrityStream   := ProcessAttribute(FILE_ATTRIBUTE_INTEGRITY_STREAM);
fAttributesDecoded.Normal            := ProcessAttribute(FILE_ATTRIBUTE_NORMAL);
fAttributesDecoded.NotContentIndexed := ProcessAttribute(FILE_ATTRIBUTE_NOT_CONTENT_INDEXED);
fAttributesDecoded.NoScrubData       := ProcessAttribute(FILE_ATTRIBUTE_NO_SCRUB_DATA);
fAttributesDecoded.Offline           := ProcessAttribute(FILE_ATTRIBUTE_OFFLINE);
fAttributesDecoded.ReadOnly          := ProcessAttribute(FILE_ATTRIBUTE_READONLY);
fAttributesDecoded.ReparsePoint      := ProcessAttribute(FILE_ATTRIBUTE_REPARSE_POINT);
fAttributesDecoded.SparseFile        := ProcessAttribute(FILE_ATTRIBUTE_SPARSE_FILE);
fAttributesDecoded.System            := ProcessAttribute(FILE_ATTRIBUTE_SYSTEM);
fAttributesDecoded.Temporary         := ProcessAttribute(FILE_ATTRIBUTE_TEMPORARY);
fAttributesDecoded.Virtual           := ProcessAttribute(FILE_ATTRIBUTE_VIRTUAL);
end;

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
procedure TWinFileInfo.LoadVersionInfo;
var
  Dummy:  DWORD;
begin
fVerInfoSize := GetFileVersionInfoSize(PChar(StrToWin(fLongName)),Dummy);
fVersionInfoPresent := fVerInfoSize > 0;
If fVersionInfoPresent then
  begin
    fVerInfoData := AllocMem(fVerInfoSize);
    If GetFileVersionInfo(PChar(StrToWin(fLongName)),0,fVerInfoSize,fVerInfoData) then
      begin
        VersionInfo_LoadTranslations;
        // parsing must be done here, before strings loading
        If lsaParseVersionInfo in fLoadingStrategy then
          begin
            VersionInfo_Parse;
            If lsaVerInfoExtractTranslations in fLoadingStrategy then
              VersionInfo_ExtractTranslations;
            VersionInfo_EnumerateKeys;
          end;
        VersionInfo_LoadStrings;
      end
    else
      begin
        FreeMem(fVerInfoData,fVerInfoSize);
        fVerInfoData := nil;
        fVerInfoSize := 0;
        fVersionInfoPresent := False;
      end;
  end;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

{$IFDEF FPCDWM}{$PUSH}W5057{$ENDIF}
procedure TWinFileInfo.LoadFixedFileInfo;
var
  FFIPtr:   Pointer;
  FFISize:  UInt32;
begin
fVersionInfoFFIPresent := VerQueryValue(fVerInfoData,'\',FFIPtr,FFISize);
If fVersionInfoFFIPresent then
  begin
    If FFISize = SizeOf(TVSFixedFileInfo) then
      fVersionInfoFFI := PVSFixedFileInfo(FFIPtr)^
    else
      raise EWFIException.CreateFmt('TWinFileInfo.LoadFixedFileInfo: Wrong size of fixed file information (got %d, expected %d).',[FFISize,SizeOf(TVSFixedFileInfo)]);
  end;
end;
{$IFDEF FPCDWM}{$POP}{$ENDIF}

//------------------------------------------------------------------------------

procedure TWinFileInfo.DecodeFixedFileInfo;
var
  FFIWorkFileFlags: DWORD;

  Function VersionToStr(Low,High: DWORD): String;
  begin
    Result := Format('%d.%d.%d.%d',[High shr 16,High and $FFFF,Low shr 16,Low and $FFFF]);
  end;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

  Function GetFlagText(Flag: DWORD; Data: array of TWFIFlagText; const NotFound: String): String;
  var
    i:  Integer;
  begin
    Result := NotFound;
    For i := Low(Data) to High(Data) do
      If Data[i].Flag = Flag then
        begin
          Result := Data[i].Text;
          Break;
        end;
  end;

begin
with fVersionInfoFFIDecoded,fVersionInfoFFI do
  begin
    FileVersionFull := (UInt64(dwFileVersionMS) shl 32) or UInt64(fVersionInfoFFI.dwFileVersionLS);
    FileVersionMembers.Major := dwFileVersionMS shr 16;
    FileVersionMembers.Minor := dwFileVersionMS and $FFFF;
    FileVersionMembers.Release := dwFileVersionLS shr 16;
    FileVersionMembers.Build := dwFileVersionLS and $FFFF;
    FileVersionStr := VersionToStr(dwFileVersionLS,dwFileVersionMS);
    ProductVersionFull := (UInt64(dwProductVersionMS) shl 32) or UInt64(dwProductVersionLS);
    ProductVersionMembers.Major := dwProductVersionMS shr 16;
    ProductVersionMembers.Minor := dwProductVersionMS and $FFFF;
    ProductVersionMembers.Release := dwProductVersionLS shr 16;
    ProductVersionMembers.Build := dwProductVersionLS and $FFFF;
    ProductVersionStr := VersionToStr(dwProductVersionLS,dwProductVersionMS);
    // mask flags
    FFIWorkFileFlags := dwFileFlags and dwFileFlagsMask;
    // decode masked flags
    FileFlags.Debug        := ((FFIWorkFileFlags) and VS_FF_DEBUG) <> 0;
    FileFlags.InfoInferred := ((FFIWorkFileFlags) and VS_FF_INFOINFERRED) <> 0;
    FileFlags.Patched      := ((FFIWorkFileFlags) and VS_FF_PATCHED) <> 0;
    FileFlags.Prerelease   := ((FFIWorkFileFlags) and VS_FF_PRERELEASE) <> 0;
    FileFlags.PrivateBuild := ((FFIWorkFileFlags) and VS_FF_PRIVATEBUILD) <> 0;
    FileFlags.SpecialBuild := ((FFIWorkFileFlags) and VS_FF_SPECIALBUILD) <> 0;
    FileOSStr := GetFlagText(dwFileOS,WFI_FFI_FILE_OS_STRS,'Unknown');
    FileTypeStr := GetFlagText(dwFileType,WFI_FFI_FILE_TYPE_STRS,'Unknown');
    case fVersionInfoFFI.dwFileType of
      VFT_DRV:  FileSubTypeStr := GetFlagText(dwFileType,WFI_FFI_FILE_SUBTYPE_DRV_STRS,'Unknown');
      VFT_FONT: FileSubTypeStr := GetFlagText(dwFileType,WFI_FFI_FILE_SUBTYPE_FONT_STRS,'Unknown');
      VFT_VXD:  FileSubTypeStr := IntToHex(dwFileSubtype,8);
    else
      FileSubTypeStr := '';
    end;
    FileDateFull := (UInt64(dwFileDateMS) shl 32) or UInt64(dwFileDateLS);
  end;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.Clear;
begin
// do not clear file handle, existence indication and names
fSize := 0;
fSizeStr := '';
fCreationTime := 0;
fLastAccessTime := 0;
fLastWriteTime := 0;
fNumberOfLinks := 0;
fVolumeSerialNumber := 0;
fFileID := 0;
// attributes
fAttributesFlags := 0;
fAttributesStr := '';
fAttributesText := '';
FillChar(fAttributesDecoded,SizeOf(fAttributesDecoded),0);
// version info unparsed data
If Assigned(fVerInfoData) and (fVerInfoSize <> 0) then
  FreeMem(fVerInfoData,fVerInfoSize);
fVerInfoData := nil;
fVerInfoSize := 0;
// version info
fVersionInfoPresent := False;
// fixed file info
fVersionInfoFFIPresent := False;
FillChar(fVersionInfoFFI,SizeOf(fVersionInfoFFI),0);
fVersionInfoFFIDecoded.FileVersionStr := '';
fVersionInfoFFIDecoded.ProductVersionStr := '';
fVersionInfoFFIDecoded.FileOSStr := '';
fVersionInfoFFIDecoded.FileTypeStr := '';
fVersionInfoFFIDecoded.FileSubTypeStr := '';
FillChar(fVersionInfoFFIDecoded,SizeOf(fVersionInfoFFIDecoded),0);
// version info partially parsed data
fVersionInfoStruct.Key := '';
SetLength(fVersionInfoStruct.StringFileInfos,0);
SetLength(fVersionInfoStruct.VarFileInfos,0);
FillChar(fVersionInfoStruct,SizeOf(fVersionInfoStruct),0);
{
  do not free fVersionInfoStruct.FixedFileInfo pointer, it was pointing into
  fVerInfoData memory block which was already freed
}
// version info fully parsed data
fVersionInfoParsed := False;
SetLength(fVersionInfoStringTables,0);
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.Initialize(const FileName: String);
var
  LastError:  Integer;
begin
// set file paths
{$IFDEF UTF8Wrappers}
fLongName := ExpandFileNameUTF8(FileName);
{$ELSE}
fLongName := ExpandFileName(FileName);
{$ENDIF}
SetLength(fShortName,MAX_PATH);
SetLength(fShortName,GetShortPathName(PChar(StrToWin(fLongName)),PChar(fShortName),Length(fShortName)));
fShortName := WinToStr(fShortName);
// open file and check its existence (with this arguments cannot open directory, so that one is good)
fFileHandle := CreateFile(PChar(StrToWin(fLongName)),0,FILE_SHARE_READ or FILE_SHARE_WRITE,nil,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,0);
If fFileHandle <> INVALID_HANDLE_VALUE then
  begin
    // file was successfully opened, we can assume it exists
    fExists := True;
    // and now for stages of loading and parsing
    If lsaLoadBasicInfo in fLoadingStrategy then
      begin
        LoadBasicInfo;
        If lsaDecodeBasicInfo in fLoadingStrategy then
          DecodeBasicInfo;
      end;
    If lsaLoadVersionInfo in fLoadingStrategy then
      begin
        LoadVersionInfo;  // also do parsing and other stuff
        If fVersionInfoPresent then
          begin
            // following can be done only if version info is loaded
            If lsaLoadFixedFileInfo in fLoadingStrategy then
              begin
                LoadFixedFileInfo;
                If fVersionInfoFFIPresent and (lsaDecodeFixedFileInfo in fLoadingStrategy) then
                  DecodeFixedFileInfo;
              end;
          end;
      end;
  end
else
  begin
    // CreateFile has failed
    LastError := GetLastError;
    If LastError = ERROR_FILE_NOT_FOUND then
      fExists := False
    else
      raise EWFISystemError.CreateFmt('CreateFile failed with error 0x%.8x.',[LastError]);
  end;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.Finalize;
begin
Clear;
CloseHandle(fFileHandle);
fFileHandle := INVALID_HANDLE_VALUE;
end;

{-------------------------------------------------------------------------------
    TWinFileInfo - public methods
-------------------------------------------------------------------------------}

constructor TWinFileInfo.Create(LoadingStrategy: TWFILoadingStrategy = WFI_LS_All);
var
  ModuleFileName: String;
begin
SetLength(ModuleFileName,MAX_PATH);
SetLength(ModuleFileName,GetModuleFileName(hInstance,PChar(ModuleFileName),Length(ModuleFileName)));
ModuleFileName := WinToStr(ModuleFileName);
Create(ModuleFileName,LoadingStrategy);
end;

//--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

constructor TWinFileInfo.Create(const FileName: String; LoadingStrategy: TWFILoadingStrategy = WFI_LS_All);
begin
inherited Create;
fLoadingStrategy := LoadingStrategy;
{$WARN SYMBOL_PLATFORM OFF}
{$IF not Defined(FPC) and (CompilerVersion >= 18)} // Delphi 2006+
fFormatSettings := TFormatSettings.Create(LOCALE_USER_DEFAULT);
{$ELSE}
GetLocaleFormatSettings(LOCALE_USER_DEFAULT,fFormatSettings);
{$IFEND}
{$WARN SYMBOL_PLATFORM ON}
Initialize(FileName);
end;

//------------------------------------------------------------------------------

destructor TWinFileInfo.Destroy;
begin
Finalize;
inherited;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.Refresh;
begin
Clear;
Finalize;
Initialize(fLongName);
end;
 
//------------------------------------------------------------------------------

Function TWinFileInfo.IndexOfVersionInfoStringTable(Translation: DWORD): Integer;
var
  i:  Integer;
begin
Result := -1;
For i := Low(fVersionInfoStringTables) to High(fVersionInfoStringTables) do
  If fVersionInfoStringTables[i].Translation.Translation = Translation then
    begin
      Result := i;
      Exit;
    end;
end;

//------------------------------------------------------------------------------

Function TWinFileInfo.IndexOfVersionInfoString(Table: Integer; const Key: String): Integer;
var
  i:  Integer;
begin
Result := -1;
with GetVersionInfoStringTable(Table) do
  For i := Low(Strings) to High(Strings) do
    If AnsiSameText(Strings[i].Key,Key) then
      begin
        Result := i;
        Exit;
      end;
end;

//------------------------------------------------------------------------------

procedure TWinFileInfo.CreateReport(Strings: TStrings);
var
  i,j:  Integer;
  Len:  Integer;
begin
Strings.Add('=== TWinInfoFile report, created on ' + DateTimeToStr(Now,fFormatSettings) + ' ===');
Strings.Add(sLineBreak + '--- General info ---' + sLineBreak);
Strings.Add('  Exists:     ' + BoolToStr(fExists,True));
Strings.Add('  Long name:  ' + LongName);
Strings.Add('  Short name: ' + ShortName);
Strings.Add(sLineBreak + 'Size:' + sLineBreak);
Strings.Add('  Size:    ' + IntToStr(fSize));
Strings.Add('  SizeStr: ' + SizeStr);
Strings.Add(sLineBreak + 'Time:' + sLineBreak);
Strings.Add('  Created:     ' + DateTimeToStr(fCreationTime));
Strings.Add('  Last access: ' + DateTimeToStr(fLastAccessTime));
Strings.Add('  Last write:  ' + DateTimeToStr(fLastWriteTime));
Strings.Add(sLineBreak + 'Others:' + sLineBreak);
Strings.Add('  Number of links:      ' + IntToStr(fNumberOfLinks));
Strings.Add('  Volume serial number: ' + IntToHex(fVolumeSerialNumber,8));
Strings.Add('  File ID:              ' + IntToHex(fFileID,16));
Strings.Add(sLineBreak + 'Attributes:' + sLineBreak);
Strings.Add('  Attributes flags:  ' + IntToHex(fAttributesFlags,8));
Strings.Add('  Attributes string: ' + fAttributesStr);
Strings.Add('  Attributes text:   ' + fAttributesText);
Strings.Add(sLineBreak + '  Attributes decoded:');
Strings.Add('    Archive:             ' + BoolToStr(fAttributesDecoded.Archive,True));
Strings.Add('    Compressed:          ' + BoolToStr(fAttributesDecoded.Compressed,True));
Strings.Add('    Device:              ' + BoolToStr(fAttributesDecoded.Device,True));
Strings.Add('    Directory:           ' + BoolToStr(fAttributesDecoded.Directory,True));
Strings.Add('    Encrypted:           ' + BoolToStr(fAttributesDecoded.Encrypted,True));
Strings.Add('    Hidden:              ' + BoolToStr(fAttributesDecoded.Hidden,True));
Strings.Add('    Integrity stream:    ' + BoolToStr(fAttributesDecoded.IntegrityStream,True));
Strings.Add('    Normal:              ' + BoolToStr(fAttributesDecoded.Normal,True));
Strings.Add('    Not content indexed: ' + BoolToStr(fAttributesDecoded.NotContentIndexed,True));
Strings.Add('    No scrub data:       ' + BoolToStr(fAttributesDecoded.NoScrubData,True));
Strings.Add('    Offline:             ' + BoolToStr(fAttributesDecoded.Offline,True));
Strings.Add('    Read only:           ' + BoolToStr(fAttributesDecoded.ReadOnly,True));
Strings.Add('    Reparse point:       ' + BoolToStr(fAttributesDecoded.ReparsePoint,True));
Strings.Add('    Sparse file:         ' + BoolToStr(fAttributesDecoded.SparseFile,True));
Strings.Add('    System:              ' + BoolToStr(fAttributesDecoded.System,True));
Strings.Add('    Temporary:           ' + BoolToStr(fAttributesDecoded.Temporary,True));
Strings.Add('    Vitual:              ' + BoolToStr(fAttributesDecoded.Virtual,True));
If fVersionInfoPresent then
  begin
    Strings.Add(sLineBreak + '--- File version info ---');
    If fVersionInfoFFIPresent then
      begin
        Strings.Add(sLineBreak + 'Fixed file info:' + sLineBreak);
        Strings.Add('  Signature:         ' + IntToHex(fVersionInfoFFI.dwSignature,8));
        Strings.Add('  Struct version:    ' + IntToHex(fVersionInfoFFI.dwStrucVersion,8));
        Strings.Add('  File version H:    ' + IntToHex(fVersionInfoFFI.dwFileVersionMS,8));
        Strings.Add('  File version L:    ' + IntToHex(fVersionInfoFFI.dwFileVersionLS,8));
        Strings.Add('  Product version H: ' + IntToHex(fVersionInfoFFI.dwProductVersionMS,8));
        Strings.Add('  Product version L: ' + IntToHex(fVersionInfoFFI.dwProductVersionLS,8));
        Strings.Add('  File flags mask :  ' + IntToHex(fVersionInfoFFI.dwFileFlagsMask,8));
        Strings.Add('  File flags:        ' + IntToHex(fVersionInfoFFI.dwFileFlags,8));
        Strings.Add('  File OS:           ' + IntToHex(fVersionInfoFFI.dwFileOS,8));
        Strings.Add('  File type:         ' + IntToHex(fVersionInfoFFI.dwFileType,8));
        Strings.Add('  File subtype:      ' + IntToHex(fVersionInfoFFI.dwFileSubtype,8));
        Strings.Add('  File date H:       ' + IntToHex(fVersionInfoFFI.dwFileDateMS,8));
        Strings.Add('  File date L:       ' + IntToHex(fVersionInfoFFI.dwFileDateLS,8));
        Strings.Add(sLineBreak + '  Fixed file info decoded:');
        Strings.Add('    File version full:      ' + IntToHex(fVersionInfoFFIDecoded.FileVersionFull,16));
        Strings.Add('    File version members:');
        Strings.Add('      Major:   ' + IntToStr(fVersionInfoFFIDecoded.FileVersionMembers.Major));
        Strings.Add('      Minor:   ' + IntToStr(fVersionInfoFFIDecoded.FileVersionMembers.Minor));
        Strings.Add('      Release: ' + IntToStr(fVersionInfoFFIDecoded.FileVersionMembers.Release));
        Strings.Add('      Build:   ' + IntToStr(fVersionInfoFFIDecoded.FileVersionMembers.Build));
        Strings.Add('    File version string:    ' + fVersionInfoFFIDecoded.FileVersionStr);
        Strings.Add('    Product version full:   ' + IntToHex(fVersionInfoFFIDecoded.ProductVersionFull,16));
        Strings.Add('    Product version members:');
        Strings.Add('      Major:   ' + IntToStr(fVersionInfoFFIDecoded.ProductVersionMembers.Major));
        Strings.Add('      Minor:   ' + IntToStr(fVersionInfoFFIDecoded.ProductVersionMembers.Minor));
        Strings.Add('      Release: ' + IntToStr(fVersionInfoFFIDecoded.ProductVersionMembers.Release));
        Strings.Add('      Build:   ' + IntToStr(fVersionInfoFFIDecoded.ProductVersionMembers.Build));
        Strings.Add('    Product version string: ' + fVersionInfoFFIDecoded.ProductVersionStr);
        Strings.Add('    File flags:');
        Strings.Add('      Debug:         ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.Debug,True));
        Strings.Add('      Info inferred: ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.InfoInferred,True));
        Strings.Add('      Patched:       ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.Patched,True));
        Strings.Add('      Prerelease:    ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.Prerelease,True));
        Strings.Add('      Private build: ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.PrivateBuild,True));
        Strings.Add('      Special build: ' + BoolToStr(fVersionInfoFFIDecoded.FileFlags.SpecialBuild,True));
        Strings.Add('    File OS string:         ' + fVersionInfoFFIDecoded.FileOSStr);
        Strings.Add('    File type string:       ' + fVersionInfoFFIDecoded.FileTypeStr);
        Strings.Add('    File subtype string:    ' + fVersionInfoFFIDecoded.FileSubTypeStr);
        Strings.Add('    File date full:         ' + IntToHex(fVersionInfoFFIDecoded.FileDateFull,16));
      end
    else Strings.Add(sLineBreak + 'Fixed file info not present.');
    If VersionInfoTranslationCount > 0 then
      begin
        Strings.Add(sLineBreak + 'Version info translations:' + sLineBreak);
        Strings.Add('  Translation count: ' + IntToStr(VersionInfoTranslationCount));
        For i := 0 to Pred(VersionInfoTranslationCount) do
          begin
            Strings.Add(sLineBreak + Format('  Translation %d:',[i]));
            Strings.Add('    Language:        ' + IntToStr(VersionInfoTranslations[i].Language));
            Strings.Add('    Code page:       ' + IntToStr(VersionInfoTranslations[i].CodePage));
            Strings.Add('    Translation:     ' + IntToHex(VersionInfoTranslations[i].Translation,8));
            Strings.Add('    Language string: ' + VersionInfoTranslations[i].LanguageStr);
            Strings.Add('    Language name:   ' + VersionInfoTranslations[i].LanguageName);
          end;
        end
      else Strings.Add(sLineBreak + 'No translation found.');
    If VersionInfoStringTableCount > 0 then
      begin
        Strings.Add(sLineBreak + 'Version info string tables:' + sLineBreak);
        Strings.Add('  String table count: ' + IntToStr(VersionInfoStringTableCount));
        For i := 0 to Pred(VersionInfoStringTableCount) do
          begin
            Strings.Add(sLineBreak + Format('  String table %d (%s):',[i,fVersionInfoStringTables[i].Translation.LanguageName]));
            Len := 0;
            For j := 0 to Pred(VersionInfoStringCount[i]) do
              If Length(VersionInfoStrings[i,j].Key) > Len then Len := Length(VersionInfoStrings[i,j].Key);
            For j := 0 to Pred(VersionInfoStringCount[i]) do
              Strings.Add(Format('    %s: %s%s',[VersionInfoStrings[i,j].Key,
                StringOfChar(' ',Len - Length(VersionInfoStrings[i,j].Key)),
                VersionInfoStrings[i,j].Value]));
          end;
        end
      else Strings.Add(sLineBreak + 'No string table found.');
  end
else Strings.Add(sLineBreak + 'File version information not present.');
end;

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Function TWinFileInfo.CreateReport: String;
var
  Strings:  TStringList;
begin
Strings := TStringList.Create;
try
  CreateReport(Strings);
  Result := Strings.Text;
finally
  Strings.Free;
end;
end;

{$ENDIF LimitedImplementation}

end.
