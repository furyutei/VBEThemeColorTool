# -*- coding: utf-8 -*-

'''
VBEThemeColorTool (vbetctool)
===
This is a command line tool to apply a theme for [VBEThemeColorEditor]](https://github.com/gallaux/VBEThemeColorEditor) to VBE7.DLL.  

'''

__version__ = '0.1.2'
__author__ = 'furyu'
__author_email__ = 'furyutei@gmail.com'
__copyright__ = 'Copyright (c) 2021 furyu'
__license__ = 'MIT'
__url__ = 'https://github.com/furyutei/VBEThemeColorTool'


import sys
try:
  sys.stdout.reconfigure( encoding=sys.stdout.encoding, errors='backslashreplace' )
except Exception as error:
  import io
  sys.stdout = io.TextIOWrapper( sys.stdout.buffer, encoding=sys.stdout.encoding, errors='backslashreplace' )
import os
import traceback
import re
import xmltodict
import hashlib
import winreg
from win32api import GetFileVersionInfo, LOWORD, HIWORD

DefaultMaxKeywordOffset = 512

ThemeColorsDef = dict(
  OriginalPattern = bytes.fromhex('ffffff00c0c0c0008080800000000000ff00000080000000ffff00008080000000ff00000080000000ffff00008080000000ff0000008000ff00ff0080008000'),
  PrevMark = bytes.fromhex('ffffffff'),
  NextMark = None,
  BeforeKeyword = r'Options'.encode('utf-8'),
  AfterKeyword = r'Software\Microsoft\VBA'.encode('utf-8'),
  MaxKeywordOffset = 1024,
  ColorIdSequence = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16),
)

PaletteDef = dict(
  OriginalPattern = bytes.fromhex('00000000000080000080000000808000800000008000800080800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff00'),
  PrevMark = bytes.fromhex('ffffffff'),
  NextMark = None,
  BeforeKeyword = r'VB4TestMMF'.encode('utf-8'),
  AfterKeyword = r'Software\Microsoft\Visual Basic\6.0\External Tools'.encode('utf-8'),
  ColorIdSequence = (4, 14, 10, 12, 6, 16, 8, 2, 3, 13, 9, 11, 5, 15, 7, 1),
)

VerboseOutput = False
StrictDllVerification = True


def GetFileInfo(FilePath): #{
  FixedInfo = GetFileVersionInfo(FilePath, '\\')
  ms = FixedInfo['FileVersionMS']
  ls = FixedInfo['FileVersionLS']
  Version = f'{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}.{LOWORD(ls)}'
  
  (lang, codepage) = GetFileVersionInfo(FilePath, r'\VarFileInfo\Translation')[0]
  codepage_path = fr'\StringFileInfo\{lang:04X}{codepage:04X}'
  StringInfo = {}
  for prop_name in ('Comments', 'InternalName', 'ProductName','CompanyName', 'LegalCopyright', 'ProductVersion','FileDescription', 'LegalTrademarks', 'PrivateBuild','FileVersion', 'OriginalFilename', 'SpecialBuild',):
    StringInfo[prop_name] = GetFileVersionInfo(FilePath, fr'{codepage_path}\{prop_name}')
  
  return dict(
    Version = Version,
    FixedInfo = FixedInfo,
    StringInfo = StringInfo,
  )
#} // def GetFileInfo()


def IsVbeDll(FilePath): #{
  try:
    DllFileInfo = GetFileInfo(FilePath)
    return (DllFileInfo['StringInfo']['ProductName'] == 'Visual Basic Environment')
  except Exception as error:
    return False
#} // def IsVbeDll()


def GetColorPattern(ColorMap, ColorIdSequence): #{
  return bytes.fromhex('00'.join([ColorMap[color_id]['HexColor'] for color_id in ColorIdSequence]) + '00')
#} // def GetColorPattern()


def GetThemeInfo(ThemeFile): #{
  with open(ThemeFile, 'rb') as fp:
    theme_xml = fp.read()
  theme_dict = xmltodict.parse(theme_xml)
  ThemeInfo = dict(
    name = theme_dict['VbeTheme']['@name'],
    desc = theme_dict['VbeTheme']['@desc'],
    ColorMap = {},
  )
  ColorMap = ThemeInfo[ 'ColorMap' ]
  for color_info in theme_dict['VbeTheme']['ThemeColors']['Color']:
    colorID = int(color_info['@colorID'])
    ColorMap[colorID] = dict(
      colorID = colorID,
      HexColor = color_info['@HexColor'],
    )
  ThemeInfo.update(dict(
    ThemeColorPattern = GetColorPattern(ColorMap, ThemeColorsDef['ColorIdSequence']),
    PalettePattern = GetColorPattern(ColorMap, PaletteDef['ColorIdSequence']),
  ))
  return ThemeInfo
#} // def GetThemeInfo()


def GetVbeDllBytes(VbeDllPath): #{
  with open(VbeDllPath, 'rb' ) as fp:
    VbeDllBytes = fp.read()
  return VbeDllBytes
#} // def GetVbeDllBytes()


def GetFirstPatternPosition(VbeDllBytes, PatternDef): #{
  OriginalPattern = PatternDef['OriginalPattern']
  PrevMark = PatternDef.get('PrevMark')
  if PrevMark is None: PrevMark = b''
  NextMark = PatternDef.get('NextMark')
  if NextMark is None: NextMark = b''
  FoundPosition = VbeDllBytes.find(PrevMark + OriginalPattern + NextMark)
  if FoundPosition < 0: return -1
  
  if StrictDllVerification:
    MaxKeywordOffset = PatternDef.get('MaxKeywordOffset', DefaultMaxKeywordOffset)
    BeforeKeyword = PatternDef.get('BeforeKeyword')
    if BeforeKeyword is not None:
      BeforeKeywordPosition = VbeDllBytes[:FoundPosition].rfind(BeforeKeyword)
      if BeforeKeywordPosition < 0: return -1
      BeforeKeywordOffset = FoundPosition - BeforeKeywordPosition
      if MaxKeywordOffset < BeforeKeywordOffset: return -1
    
    AfterKeyword = PatternDef.get('AfterKeyword')
    if AfterKeyword is not None:
      AfterKeywordPosition = VbeDllBytes[FoundPosition + len(PrevMark + OriginalPattern + NextMark):].find(AfterKeyword)
      if AfterKeywordPosition < 0: return -1
      AfterKeywordOffset = AfterKeywordPosition + 1
      if MaxKeywordOffset < AfterKeywordOffset: return -1
  
  return FoundPosition + len(PrevMark)
#} // def GetFirstPatternPosition()


def GetVbeDllBytesHash(VbeDllBytes): #{
  return hashlib.sha256(VbeDllBytes).hexdigest()
#} // def GetVbeDllBytesHash(


def GetVbeDllInfo(VbeDllPath): #{
  if StrictDllVerification and not IsVbeDll(VbeDllPath):
    raise Exception(f'File ("{VbeDllPath}") is not VBEx.DLL.')
  
  VbeDllBytes = GetVbeDllBytes(VbeDllPath)
  theme_colors_position = GetFirstPatternPosition(VbeDllBytes, ThemeColorsDef)
  palette_position = GetFirstPatternPosition(VbeDllBytes, PaletteDef)
  return dict(
    path = VbeDllPath,
    basename = os.path.basename(VbeDllPath),
    dirname = os.path.dirname(VbeDllPath),
    bytes = VbeDllBytes,
    hash = GetVbeDllBytesHash(VbeDllBytes),
    theme_colors_position = theme_colors_position,
    palette_position = palette_position,
    is_supported = (0 <= theme_colors_position) and (0 <= palette_position),
  )
#} // def GetVbeDllInfo()


'''
#def PatchVbeDllBytes(VbeDllBytes, ThemeInfo): #{
#  parts = VbeDllBytes.partition(ThemeColorsDef['OriginalPattern'])
#  if len(parts[1]) < 1: return None
#  parts = (parts[0] + ThemeInfo['ThemeColorPattern'] + parts[2]).partition(PaletteDef['OriginalPattern'])
#  if len(parts[1]) < 1: return None
#  return parts[0] + ThemeInfo['PalettePattern'] + parts[2]
##} // def PatchVbeDllBytes()
'''

def GetPatchedVbeDllBytes(VbeDllInfo, ThemeInfo): #{
  work_bytes = VbeDllInfo['bytes']
  work_bytes = work_bytes[:VbeDllInfo['theme_colors_position']] + ThemeInfo['ThemeColorPattern'] + work_bytes[VbeDllInfo['theme_colors_position']+len(ThemeColorsDef['OriginalPattern']):]
  work_bytes = work_bytes[:VbeDllInfo['palette_position']] + ThemeInfo['PalettePattern'] + work_bytes[VbeDllInfo['palette_position']+len(PaletteDef['OriginalPattern']):]
  return work_bytes
#} // def GetPatchedVbeDllBytes()


def NormalizeCodeColors(CodeColors): #{
  CodeColors = CodeColors.strip()
  if CodeColors == 'delete': return None
  color_list = [color_id if 0 <= color_id and color_id <= 16 else 0 for color_id in map(int, re.split('\s+', CodeColors))]
  if len(color_list) < 10 or 16 < len(color_list):
    raise Exception(f'CodeColors("{CodeColors}") error')
  color_list += [0] * (16 - len(color_list))
  return ' '.join(map(str, color_list)) + ' '
#} // def NormalizeCodeColors()


def ApplyThemeToVbeDll(VbeDllPath, ThemeFile): #{
  if not os.path.exists(ThemeFile):
    raise Exception(f'Theme file ("{ThemeFile}") is not found.')
  
  try:
    ThemeInfo = GetThemeInfo(ThemeFile)
  except Exception as error:
    raise Exception(f'Failed to parse theme file ("{ThemeFile}").')
  
  if not os.path.exists(VbeDllPath):
    raise Exception(f'VBEx.DLL file ("{VbeDllPath}") is not found.')
  
  BackupVbeDllPath = VbeDllPath + '.BAK'
  
  ExistVbeDllInfo = GetVbeDllInfo(VbeDllPath)
  ExistBackupVbeDllInfo = GetVbeDllInfo(BackupVbeDllPath) if os.path.exists(BackupVbeDllPath) else None
  
  SourceVbeDllInfo = ExistVbeDllInfo if ExistVbeDllInfo['is_supported'] else (ExistBackupVbeDllInfo if (ExistBackupVbeDllInfo is not None and ExistBackupVbeDllInfo['is_supported']) else None)
  if SourceVbeDllInfo is None:
    raise Exception(f'Specified VBEx.DLL ("{VbeDllPath}") is not applicable.')
  
  if ExistBackupVbeDllInfo is None:
    with open(BackupVbeDllPath, 'wb') as fp:
      fp.write(SourceVbeDllInfo['bytes'])
  else:
    if ExistBackupVbeDllInfo['hash'] != SourceVbeDllInfo['hash']:
      with open(BackupVbeDllPath, 'wb') as fp:
        fp.write(SourceVbeDllInfo['bytes'])
  
  PatchedVbeDllBytes = GetPatchedVbeDllBytes(SourceVbeDllInfo, ThemeInfo)
  PatchedVbeDllBytesHash = GetVbeDllBytesHash(PatchedVbeDllBytes)
  
  if PatchedVbeDllBytesHash != ExistVbeDllInfo['hash']:
    with open(VbeDllPath, 'wb') as fp:
      fp.write(PatchedVbeDllBytes)
  
  return dict(
    dll_path = VbeDllPath,
    old_dll_hash = ExistVbeDllInfo['hash'],
    new_dll_hash = PatchedVbeDllBytesHash,
    backup_path = BackupVbeDllPath,
    old_backup_hash = ExistBackupVbeDllInfo['hash'] if ExistBackupVbeDllInfo is not None else None,
    new_backup_hash = SourceVbeDllInfo['hash'],
  )
#} // def ApplyThemeToVbeDll()


def ApplyColorsToRegistry(KeyName, CodeColors): #{
  try:
    CodeColors = NormalizeCodeColors(CodeColors)
    with winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\VBA\7.1\Common', access=winreg.KEY_SET_VALUE) as key:
      if CodeColors is None:
        try:
          winreg.DeleteValue(key, KeyName)
        except FileNotFoundError as error:
          pass
      else:
        winreg.SetValueEx(key, KeyName, 0, winreg.REG_SZ, CodeColors)
    return CodeColors
  except Exception as error:
    raise Exception(f'Failed to apply specified {KeyName}("{CodeColors}").')
#} // def ApplyColorsToRegistry()


if __name__ == '__main__': #{
  import argparse
  
  argparser = argparse.ArgumentParser()
  
  argparser.add_argument(
    '-v', '--version',
    action = 'version',
    version = f'%(prog)s {__version__}'
  )
  
  argparser.add_argument(
    '-l', '--target-dll',
    type = str,
    metavar = '<VBE7.DLL>',
    help = '%(metavar)s: target DLL (VBE7.DLL) file path to apply theme',
    dest = 'VbeDllPath'
  )
  
  argparser.add_argument(
    '-t', '--theme-xml-file',
    type = str,
    metavar = '<THEME>',
    help = '%(metavar)s: VBEThemeColorEditor theme XML file to apply',
    dest = 'ThemeFile'
  )
  
  argparser.add_argument(
    '-f', '--fore-colors',
    type = str,
    metavar = '<FORE COLORS>',
    help = '%(metavar)s: foreground color values to be set in the registry ("delete": delete value)',
    dest = 'CodeForeColors'
  )
  
  argparser.add_argument(
    '-b', '--back-colors',
    type = str,
    metavar = '<BACK COLORS>',
    help = '%(metavar)s: background color values to be set in the registry ("delete": delete value)',
    dest = 'CodeBackColors'
  )
  
  argparser.add_argument(
    '-n', '--no-strict-verification',
    action = 'store_false',
    help = 'turn off strict DLL verifiction',
    dest = 'StrictDllVerification',
  )
  
  argparser.add_argument(
    '-V', '--verbose',
    action = 'store_true',
    help = 'turn on verbose output',
    dest = 'VerboseOutput',
  )
  
  args = argparser.parse_args()
  VerboseOutput = args.VerboseOutput
  StrictDllVerification = args.StrictDllVerification
  exec_count = 0
  
  if (args.VbeDllPath is not None) and (args.ThemeFile is not None):
    try:
      ResultInfo = ApplyThemeToVbeDll(args.VbeDllPath, args.ThemeFile)
      if VerboseOutput:
        print(f'{ResultInfo["dll_path"]}')
        if ResultInfo['new_dll_hash'] != ResultInfo['old_dll_hash']:
          print(f'{ResultInfo["old_dll_hash"]} => {ResultInfo["new_dll_hash"]}')
        else:
          print(f'{ResultInfo["new_dll_hash"]} (No change)')
        print(f'')
        print(f'{ResultInfo["backup_path"]}')
        if ResultInfo["old_backup_hash"] is None:
          print(f'{ResultInfo["new_backup_hash"]} (Created)')
        elif ResultInfo['new_backup_hash'] != ResultInfo['old_backup_hash']:
          print(f'{ResultInfo["old_backup_hash"]} => {ResultInfo["new_backup_hash"]}')
        else:
          print(f'{ResultInfo["new_backup_hash"]} (No change)')
        print(f'')
    except Exception as error:
      print(traceback.format_exception_only(type(error), error)[0], file=sys.stderr)
      sys.exit( 1 )
    exec_count += 1
  elif (args.VbeDllPath is not None) or (args.ThemeFile is not None):
    print('Specify both "-l" and "-t" options in order to apply VBEThemeColorEditor theme.', file=sys.stderr)
    sys.exit( 1 )
  
  for KeyName in ('CodeForeColors', 'CodeBackColors',):
    CodeColors = getattr(args, KeyName, None)
    if CodeColors is None: continue
    try:
      SetCodeColors = ApplyColorsToRegistry(KeyName, CodeColors)
      if VerboseOutput: print(f'Set value of {KeyName} to "{SetCodeColors}"') if SetCodeColors is not None else print(f'Delete {KeyName} value')
    except Exception as error:
      print(traceback.format_exception_only(type(error), error)[0], file=sys.stderr)
      sys.exit( 1 )
    exec_count += 1
  
  if exec_count < 1:
    argparser.print_help()
  
  sys.exit( 0 )
#} if __name__ == '__main__'
