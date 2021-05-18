# -*- coding: utf-8 -*-

'''
VBEThemeColorTool (vbetctool)
===
This is a command line tool to apply a theme for [VBEThemeColorEditor]](https://github.com/gallaux/VBEThemeColorEditor) to VBE7.DLL.  

'''

__author__ = 'furyu (furyutei@gmail.com)'
__version__ = '0.1.1'
__copyright__ = 'Copyright (c) 2021 furyu'
__license__ = 'The MIT License (MIT)'


import sys
try:
  sys.stdout.reconfigure( encoding=sys.stdout.encoding, errors='backslashreplace' )
except Exception as error:
  import io
  sys.stdout = io.TextIOWrapper( sys.stdout.buffer, encoding=sys.stdout.encoding, errors='backslashreplace' )
import os
import re
import xmltodict
import winreg
import traceback


ThemeColorsDef = dict(
  OriginalPattern = bytes.fromhex('ffffff00c0c0c0008080800000000000ff00000080000000ffff00008080000000ff00000080000000ffff00008080000000ff0000008000ff00ff0080008000'),
  ColorIdSequence = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16),
)

PaletteDef = dict(
  OriginalPattern = bytes.fromhex('00000000000080000080000000808000800000008000800080800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff00'),
  ColorIdSequence = (4, 14, 10, 12, 6, 16, 8, 2, 3, 13, 9, 11, 5, 15, 7, 1),
)


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


def GetVbe7DllBytes(Vbe7DllPath): #{
  with open(Vbe7DllPath, 'rb' ) as fp:
    Vbe7DllBytes = fp.read()
  return Vbe7DllBytes
#} // def GetVbe7DllBytes()


def PatchVbe7DllBytes(Vbe7DllBytes, ThemeInfo): #{
  parts = Vbe7DllBytes.partition(ThemeColorsDef['OriginalPattern'])
  if len(parts[1]) < 1: return None
  parts = (parts[0] + ThemeInfo['ThemeColorPattern'] + parts[2]).partition(PaletteDef['OriginalPattern'])
  if len(parts[1]) < 1: return None
  return parts[0] + ThemeInfo['PalettePattern'] + parts[2]
#} // def PatchVbe7DllBytes()


def NormalizeCodeColors(CodeColors): #{
  CodeColors = CodeColors.strip()
  if CodeColors == 'delete': return None
  color_list = [color_id if 0 <= color_id and color_id <= 16 else 0 for color_id in map(int, re.split('\s+', CodeColors))]
  color_list += [0] * (16 - len(color_list))
  return ' '.join(map(str, color_list)) + ' '
#} // def NormalizeCodeColors(CodeColors)


def ApplyThemeToVbe7Dll(Vbe7DllPath, ThemeFile): #{
  if not os.path.exists(ThemeFile):
    raise Exception(f'Theme file ("{ThemeFile}") is not found.')
  
  if not os.path.exists(Vbe7DllPath):
    raise Exception(f'VBE7.DLL ("{Vbe7DllPath}") is not found.')
  
  TryVbe7DllPathList = [Vbe7DllPath]
  
  BackupVbe7DllPath = Vbe7DllPath + '.BAK'
  if os.path.exists(BackupVbe7DllPath):
    TryVbe7DllPathList.insert(0, BackupVbe7DllPath)
  
  try:
    ThemeInfo = GetThemeInfo(ThemeFile)
  except Exception as error:
    raise Exception(f'Failed to parse theme file ("{ThemeFile}").')
  
  for TryVbe7DllPath in TryVbe7DllPathList:
    Vbe7DllBytes = GetVbe7DllBytes(TryVbe7DllPath)
    PatchedVbe7DllBytes = PatchVbe7DllBytes(Vbe7DllBytes, ThemeInfo)
    if PatchedVbe7DllBytes is not None:
      break
  
  if PatchedVbe7DllBytes is None:
    raise Exception(f'Specified VBE7.DLL ("{Vbe7DllPath}") is not applicable.')
  
  if TryVbe7DllPath != BackupVbe7DllPath:
    with open(BackupVbe7DllPath, 'wb') as fp:
      fp.write(Vbe7DllBytes)
  
  with open(Vbe7DllPath, 'wb') as fp:
    fp.write(PatchedVbe7DllBytes)
#} // def ApplyThemeToVbe7Dll()


def ApplyColorsToRegistry(KeyName, CodeColors): #{
  try:
    CodeColors = NormalizeCodeColors(CodeColors)
    with winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\VBA\7.1\Common', access=winreg.KEY_SET_VALUE) as key:
      if CodeColors is None:
        winreg.DeleteValue(key, KeyName)
      else:
        winreg.SetValueEx(key, KeyName, 0, winreg.REG_SZ, CodeColors)
  except Exception as error:
    raise Exception(f'Failed to apply specified {KeyName}("{CodeColors}").')
#} // def main()

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
    help = '%(metavar)s: target DLL file (VBE7.DLL) path',
    dest = 'Vbe7DllPath'
  )
  
  argparser.add_argument(
    '-t', '--theme-xml-file',
    type = str,
    metavar = '<THEME>',
    help = '%(metavar)s: target theme XML file for VBEThemeColorEditor',
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
  
  args = argparser.parse_args()
  
  exec_count = 0
  if (args.Vbe7DllPath is not None) and (args.ThemeFile is not None):
    try:
      ApplyThemeToVbe7Dll(args.Vbe7DllPath, args.ThemeFile)
    except Exception as error:
      print(traceback.format_exception_only(type(error), error)[0], file=sys.stderr)
      sys.exit( 1 )
    exec_count += 1
  
  if args.CodeForeColors is not None:
    try:
      ApplyColorsToRegistry('CodeForeColors', args.CodeForeColors)
    except Exception as error:
      print(traceback.format_exception_only(type(error), error)[0], file=sys.stderr)
      sys.exit( 1 )
    exec_count += 1
  
  if args.CodeBackColors is not None:
    try:
      ApplyColorsToRegistry('CodeBackColors', args.CodeBackColors)
    except Exception as error:
      print(traceback.format_exception_only(type(error), error)[0], file=sys.stderr)
      sys.exit( 1 )
    exec_count += 1
  
  if exec_count < 1:
    argparser.print_help()
  
  sys.exit( 0 )
