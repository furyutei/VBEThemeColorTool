[VBEThemeColorTool](https://github.com/furyutei/VBEThemeColorTool)
===

- License: The MIT license  
- Copyright (c) 2021 風柳(furyu) 
- 動作確認環境: Windows 10 Pro／Excel for Microsoft 365 MSO （32 ビット／64ビット）

■ これは何？
---
[VBEThemeColorEditor](https://github.com/gallaux/VBEThemeColorEditor)を適用したVBE7.DLLを使用すると、[VBEのツール(T)→オプション(O)にて「エディターの設定」タブを開いた際に64ビット版エクセルだと異常終了する等の不具合が発生したりする模様です](https://github.com/gallaux/VBEThemeColorEditor/issues/11)（2021/05/18現在）。  
また、[32ビット版だと異常終了はしないものの、同タブからの色の設定がうまくできません](https://twitter.com/furyutei/status/1394604117788565507)。  

どうやらVBE7.DLLに対するパッチの当て方に問題があるようなので、とりあえずテーマファイルをVBE7.DLLへ適用する処理だけ抜き出してコマンド化してみました。  
テーマファイルの作成に関しては、VBEThemeColorEditorをそのままお使いください。  

■ 使い方
---
⚠ あらかじめVBE7.DLLのある場所を探し、必ず別の場所にバックアップを取っておくようにしてください。  
※ 参考までに、当方の環境だと以下の場所にありました。  

- [32ビット版] `C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1`
- [64ビット版] `C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1`

1. distフォルダにある「vbetctool.exe」を、VBEThemeColorEditorと同じディレクトリにコピーします。
1. 起動しているエクセル等のオフィスアプリを全て終了します。  
1. コマンドプロンプトを管理者権限で起動し、1. のディレクトリに移動して
    ```bat
    vbetctool -l <VBE7.DLLのフルパス> -t <適用したいテーマファイルのパス> -f <前景色の割当> -b <背景色の割当> -V
    ```
    のようにします(-Vを付加すると適用結果が表示されます)。  
    具体例を挙げると、  
    ```
    vbetctool -l "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" -t ".\Themes\VS2017 Dark.xml" -f "14 9 12 9 5 4 2 5 5 5" -b "1 6 1 4 10 1 1 1 1 6" -V
    ```
    のような感じになります。  
    ※「-l と -t」、「-f」、「-b」の各オプションはそれぞれ単独で使用できます（-l と -t については同時に指定してください）  
    ※その他オプションについては「--help」にてご確認ください  

■ 注意事項
---
- テーマを適用する際には、オリジナルのVBE7.DLLは同じフォルダ内にVBE7.DLL.BAKという名前でコピーされ、以降はこれをオリジナルとみなして使用します（VBEThemeColorEditorに準じた動作となります。VBEThemeColorEditor等により既にVBE7.DLL.BAKが存在する場合にはこれをオリジナルとみなします）。  
- オフィスアプリの更新等に伴ってVBE7.DLLが上書きされてしまい、テーマがデフォルト状態に戻ってしまう場合があります。この場合にはオフィスアプリを全て終了したうえでテーマを再適用する必要があります（「-l」と「-t」オプションで指定）。    
  このときvbetctoolはVBE7.DLLがオリジナルかどうかチェックをした上で、  
  
    1. 新しいVBE7.DLLをVBE7.DLL.BAKにコピー
    1. 新しいVBE7.DLLに対してテーマを適用
    
  という動作を行いますが、オリジナルかどうかのチェックは不完全なため、適用前に新しいVBE7.DLLを別ディレクトリにバックアップしておくことをおすすめします。  
- VBE6.DLLについては手元で確認できないため未対応とします。

■ テーマについて
---
dist/Themesフォルダにあるテーマファイルは、以下のものをお借りしております。  
- 「Default VBE.xml」、「VS2012 Dark.xml」→[VBEThemeColorEditorのものと同じ内容です](https://github.com/gallaux/VBEThemeColorEditor/tree/master/VBEThemeColorEditor/VBEThemeColorEditor/Themes)
- 「VS2017 Dark.xml」→[@mochimo様の記事中のものと同じ内容です](https://qiita.com/mochimo/items/e9be36619a76e15bc898#2-%E3%83%86%E3%83%BC%E3%83%9Exml%E3%82%92%E4%BD%9C%E6%88%90)


■ 免責事項
---
ご利用の際には全て自己責任でお願いします。  
不具合があったり、使用した結果等により万一何らかの損害を被ったりした場合でも、作者は一切関知いたしません。
