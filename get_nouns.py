"""文書ファイル（テキスト、Excel）から名詞一覧を表示
"""
import sys
import re
from pathlib import Path
from janome.tokenizer import Tokenizer
import openpyxl


DICT_CSV = './user_dict.csv'    # ユーザー辞書
POS1_TYPE = ['一般', '固有名詞', '*']


def get_part_of_speech(token, pos_type=0):
    """トークンから品詞情報を取得
    pos_type=0: 品詞
    pos_type=1: 品詞細分類1
    pos_type=2: 品詞細分類2
    """
    return token.part_of_speech.split(',')[pos_type]


def is_noun(token):
    """トークンが名詞（あるいはカスタム名詞）かを判定
    """
    pos0 = get_part_of_speech(token)
    pos1 = get_part_of_speech(token, 1)
    return pos0 in ['名詞', u'カスタム名詞'] and pos1 in POS1_TYPE


def is_text_file(filepath):
    """ファイルがテキストファイルか判定
    """
    return filepath.suffix.lower() == '.txt'


def is_excel_file(filepath):
    """ファイルがExcelファイルか判定
    """
    return filepath.suffix.lower() == '.xlsx'


def entry_nouns(noun_map, tokens):
    """トークンリストから名詞トークンを名詞mapに登録する
    """
    for token in tokens:
        if is_noun(token) and token.surface not in noun_map:
            noun_map[token.surface] = token


def get_nouns_from_text_file(filename, tokenizer):
    """テキストファイルから名詞一覧を返す
    """
    noun_map = {}
    with open(filename) as f:
        for line in f.readlines():
            entry_nouns(noun_map, tokenizer.tokenize(line))

    return noun_map


def get_cell_texts(sheet):
    """シートからセルのテキストリストを取り出す
    """
    texts = []
    for row in [sheet[i] for i in range(1, sheet.max_row+1)]:
        # 行単位で取り出し
        for cell in row:
            # セル単位で取り出し
            if cell.value is not None and isinstance(cell.value, str):
                texts.append(cell.value)
    return texts


def get_nouns_from_excel_file(filename, tokenizer):
    """Excelファイルから名詞一覧を返す
    """
    noun_map = {}
    wb = openpyxl.load_workbook(filename)
    for sheet_name in wb.sheetnames:
        # シート名も対象
        entry_nouns(noun_map, tokenizer.tokenize(sheet_name))

        # シートの内容
        for cell_text in get_cell_texts(wb[sheet_name]):
            entry_nouns(noun_map, tokenizer.tokenize(cell_text))

    return noun_map


def main(filename):
    in_file = Path(filename)
    if not in_file.exists():
        print(f'{in_file}: No such file or directory')
        return

    # トーカナイザー生成
    tokenizer = None
    if Path(DICT_CSV).exists():
        tokenizer = Tokenizer(DICT_CSV,
                              udic_type='simpledic', udic_enc='utf8')
    else:
        tokenizer = Tokenizer()

    # 名詞一覧を取得
    noun_map = {}
    if is_text_file(in_file):
        noun_map = get_nouns_from_text_file(in_file, tokenizer)
    elif is_excel_file(in_file):
        noun_map = get_nouns_from_excel_file(in_file, tokenizer)

    # 名詞一覧を表示
    for token in noun_map.values():
        # XXX	名詞,固有名詞,地域,一般,*,*,XXX,YYY,YYY
        # とある場合、
        # XXX	名詞,固有名詞,地域,一般
        # のみを表示する
        msg = re.sub(r'^(.+\t)([^,]+),([^,]+),([^,]+),([^,]+),(.*$)',
                     r'\1\2,\3,\4,\5', str(token))

        print(msg)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(f'Usage: python {sys.argv[0]} <text-file> | <excel-file>')
        exit()

    main(sys.argv[1])
