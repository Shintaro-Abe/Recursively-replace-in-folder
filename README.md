# **■1.概要**

## **(1)目的**

フォルダ内にあるファイルの文字列を一括置換。

## **(2)動作環境**

- Windows11
- Python 3.12.x
- 以下のライブラリがインストールされていること

`openpyxl`  `pptx`  `docx`

ライブラリがインストールされていない場合は、以下のコマンドを実行。

```jsx
pip install openpyxl
```

```jsx
pip install python-pptx
```

```jsx
pip install python-docx
```

## (3)対象ファイル

- Word
- Excel
- PowerPoint
- テキストファイル(.txt .yml .rb…)
- 拡張子がないテキストファイル

# ■2.使用方法

(1)replace.pyを実行。実行後、対話式で値を入力。

```jsx
python replace.py
```

(2)「フォルダのパス: 」と出力されるので、置換したいファイルが格納されているフォルダのパスを入力してEnterキーを打鍵。

```jsx
フォルダのパス: C:Path\to\folder
```

(3)「置換前: 」と出力されるので、差し替え前の文字列を入力してEnterキーを打鍵。

```jsx
置換前: String(before)
```

(4)「置換後: 」と出力されるので、差し替え後の文字列を入力してEnterキーを打鍵。

```jsx
置換後: String(after)
```

(5)置換したファイル数と置換数を出力し、処理が完了。

```jsx
nファイル、n箇所を置換完了
```