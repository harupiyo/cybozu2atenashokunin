cybozu2atenashokunin
====================

サイボウズ(R) Office のアドレス帳からエクスポートしたCSV を
宛名職人でインポートできる形に整形するツール

使い方
------

1. サイボウズからエクスポートしたCSV をaddressperson.csv の
   ファイル名でこのスクリプトと同じフォルダに用意
[TODO] utf-8? or SJIS?

2. スクリプトの実行
   $ perl convert.pl --atenashokunin > atena.csv
[TODO] utf-8? or SJIS?

3. 生成されたatena.csv を宛名職人にインポート
[TODO] 動作確認まだ


