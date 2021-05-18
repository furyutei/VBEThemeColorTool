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

1. distフォルダにある「vbetctool.exe」を、VBEThemeColorEditorと同じディレクトリにコピーします。
1. 起動しているエクセル等のオフィスアプリを全て終了します。  
1. コマンドプロンプトを管理者権限で起動し、1. のディレクトリに移動して
    ```bat
    vbetctool -l <VBE7.DLLのフルパス> -t <指定したいテーマファイルのフルパス> -f <前景色の割当> -b <背景色の割当>
    ```
    のようにします。  
    具体例を挙げると、  
    ```
    vbetctool -l "C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" -t ".\Themes\VS2017 Dark.xml" -f "14 9 12 9 5 4 2 5 5 5" -b "1 6 1 4 10 1 1 1 1 6"
    ```
    のような感じになります。  
    ※「-l と -t」、「-f」、「-b」の各オプションはそれぞれ単独で使用できます（-l と -t は同時に指定してください）  

■ 注意事項
- オリジナルのVBE7.DLLは同じフォルダ内にVBE7.DLL.BAKという名前でコピーされ、以降はこれをオリジナルとみなして使用します（VBEThemeColorEditorにより既にVBE7.DLL.BAKが存在する場合もこれをオリジナルとみなします）。  
- エクセルの更新等に伴ってVBE7.DLLが上書きされてしまい、テーマがデフォルト状態に戻ってしまう場合があります。その場合にはvbetctoolで再適用する前に、オフィスアプリをいったん全て終了した状態で

    1. 新しいVBE7.DLLを別のディレクトリにバックアップ
    2. VBE7.DLL.BAKを削除

  しておくことをおすすめします。  
- VBE6.DLLについては手元で確認できないため未対応とします。

■ テーマについて
---
dist/Themesフォルダにあるテーマファイルは、以下のものをお借りしております。  
- 「Default VBE.xml」、「VS2012 Dark.xml」→[VBEThemeColorEditorのものと同じ内容です](https://github.com/gallaux/VBEThemeColorEditor/tree/master/VBEThemeColorEditor/VBEThemeColorEditor/Themes)
- 「VS2017 Dark.xml」→[@mochimo様の記事中のものと同じ内容です](https://qiita.com/mochimo/items/e9be36619a76e15bc898#2-%E3%83%86%E3%83%BC%E3%83%9Exml%E3%82%92%E4%BD%9C%E6%88%90)


■ 免責事項
---
- ご利用の際には全て自己責任でお願いします。  
  不具合があったり、使用した結果等により万一何らかの損害を被ったりした場合でも、作者は一切関知いたしません。
