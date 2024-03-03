## TA のための自動採点システム

(以下、注意事項)

- mac:`moodle-auto.py`,
- html のタグが異なる場合は、開発者環境で正しいタグを確認してください。
- 毎回のスレッドの名前や遅延提出者の確認は、毎回手動で行ってください。
- 関数が入れ子構造なので引数多い

### ログイン用のアカウント情報

ディレクトリ直下で.env ファイルを作成し、以下のように記述してください。(xxxxxx は自分のアカウント情報)

```
STUDENT_ID=MXXXXX
PASSWORD=XXXXXX
```

ローカルであれば、student_id と password はハードコーディングでも大丈夫です。windows だと.env ファイルがうまく読み込めなかった。
ディレクトリ直下で excel フォルダを作成し、そこに受け取った名簿を記入したエクセルファイルを置いてください。先輩に聞いてもいいかも
sample.xlsx を参考にして名簿を作成すると for 文の変更などをしなくても上手く動くと思います。

---

### 環境構築

mac の人は以下の URL を参照しながら
`https://note.com/taro_98/n/n02cc073a40c8 `

1. git clone
2. pip のインストール `https://pypi.org/`
3. `pip install pipenv`
4. `pipenv shell`

上記で、pipfile.lock と同様の環境を構築できます。
(※)PC にインストールされているツールによっては、少し手間取るかもしれません。

---

### 手順

- .excel/に受け取った名簿を記入したエクセルファイルを置く
- python ファイルの TODO の部分を変更する

---

## うまく起動しない場合

- 既存の環境を削除し，作成し直す

```
pipenv --rm
pipenv --python 3.10.2
pipenv install
```
