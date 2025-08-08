import sys
import os
import csv
import xlwings as xw
import re
import shutil
import io
import MeCab
import re

ARGS = {
    "GetSettings:in_ConfigFile": "設定ファイルパス",
    "GetSettings:in_ConfigSheets": '設定ファイルのシート名を文字列の配列で指定 例：{"Sheet1","Sheet2","Sheet3"}',
    "GetSettings:out_Config": "Dictionaryクラスを指定 Dictionary(String, object)",
    "GetAkikuraData:strAkikuraFile": "商蔵データファイルパス",
    "GetAkikuraData:strSheetName": "商蔵データファイルのシート名",
    "GetAkikuraData:dtAd": "商蔵取得データのタイトル",
    "GetAkikuraData:dtAt": "商蔵取得データ",
    "GetAkikuraData:blnAExists": "商蔵データファイル存在確認結果(有: True / 無: False)",
    "LoadKanjoJournal:strImportLog": "勘定取込ログファイルパス",
    "LoadKanjoJournal:strImportFile": "勘定取込ファイルパス",
    "LoadKanjoJournal:dcSetting": "取得設定情報",
    "LoadKanjoJournal:strLogFile": "処理ログファイルパス",
    }

DESCRIPTIONS ={
    "AcMonthly04:Main":"新資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly04B:Main":"新資材等調査票(東京)データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly05:Main":"糖化資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly05B:Main":"糖化資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthlyDataOutput:Main":"商蔵奉行,勘定奉行,日コンシステムのそれぞれのシステムより経理使用データファイルを出力する",
    "GetSettings":"Excelファイルの指定したシートにある設定情報を読込み、設定名をキーとした設定情報を返す",
    "GetAkikuraData":"商蔵データファイルから該当する商蔵データシートの商蔵タイトル情報と商蔵データを取得する",
    "LoadKanjoJournal":"勘定奉行を起動し、勘定取込ファイルを取り込みする"
}

DESCRIPTIONS ={
    "AcMonthly04:Main":"新資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly04B:Main":"新資材等調査票(東京)データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly05:Main":"糖化資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthly05B:Main":"糖化資材等調査票データを作成し、勘定奉行取込ファイルを作成する",
    "AcMonthlyDataOutput:Main":"商蔵奉行,勘定奉行,日コンシステムのそれぞれのシステムより経理使用データファイルを出力する",
    "GetSettings":"Excelファイルの指定したシートにある設定情報を読込み、設定名をキーとした設定情報を返す",
    "GetAkikuraData":"商蔵データファイルから該当する商蔵データシートの商蔵タイトル情報と商蔵データを取得する",
    "LoadKanjoJournal":"勘定奉行を起動し、勘定取込ファイルを取り込みする"
}

OTHER_SHEETS = {
    "AcMonthly04": {"AcMonthly04-files":"入出力ファイル一覧",
                    "AcMonthly04-settings":"設定一覧"},
    "AcMonthly04B": {"AcMonthly04B-files":"入出力ファイル一覧",
                     "AcMonthly04B-settings":"設定一覧"},
    "AcMonthly05": {"AcMonthly05-files":"入出力ファイル一覧",
                    "AcMonthly05-settings":"設定一覧"},
    "AcMonthly05B": {"AcMonthly05B-files":"入出力ファイル一覧",
                     "AcMonthly05B-settings":"設定一覧"},
    "AcMonthlyDataOutput": {"AcMonthlyDataOutput-files":"入出力ファイル一覧",
                            "AcMonthlyDataOutput-settings":"設定一覧"}
}


class XAMLCsvReader:
    def __init__(self, csv_file):
        self.csv_file = csv_file

    def parse_csv_line(sekf, line, delimiter=','):
        # 文字列をファイルのように扱う
        f = io.StringIO(line)
        reader = csv.reader(f, delimiter=delimiter)
        return next(reader)

    def read(self):
        count = 0
        try:
            with open(self.csv_file, encoding='utf-8') as f:
                count += 1
                lines = {"args":[], "flow":[]}
                section = "args"
                buffer = ""
                line = ""
                for line in f:
                    row = line.strip()
                    if row == "\ufeffArguments" or row == "Name,Type"  or row == "":
                        continue
                    elif row == "DisplayName,Path,Type,Properties" or row == "\ufeffDisplayName,Path,Type,Properties":
                        section = "flow"
                        continue
                    else:
                        if row.find("<NewDataSet>") >= 0 and not row.endswith('</NewDataSet>"'):
                            buffer = row
                            continue
                        elif len(buffer) > 0 and not row.endswith('</NewDataSet>"'):
                            buffer += row
                            continue
                        elif len(buffer) > 0 and row.endswith('</NewDataSet>"'):
                            buffer += row
                            row = buffer
                            buffer = ""
                            print(f"buffer:{row}")

                    lines[section].append([item if item != "" else "-" for item in self.parse_csv_line(row)])
        except Exception as e:
            print(f"Error reading CSV file {self.csv_file}\n{count}:{line}\nError:{e}")
        
        return lines

def add_args(app, rows, new_sheet, sheetKey):
    start_row = 10
    for i, row in enumerate(rows):
        insert_row = target_row = start_row + i
        if i > 0:
            new_sheet.range(f"A{insert_row}:BT{insert_row}").select()
            app.selection.insert(shift='down')  # 行を挿入

        new_sheet.range(f"A{start_row}:BT{start_row}").select()
        app.api.Selection.Copy()  # 行をコピー
        new_sheet.range(f"A{target_row}:BT{target_row}").select()
        app.selection.paste # xlPasteAll
        new_sheet.range(f"A{target_row}:BT{target_row}").row_height = new_sheet.range(f"A{start_row}:BO{start_row}").row_height  # 行の高さをコピー元と同じにする


        # データ書き込み
        direction = "In/Out"
        if row[1].startswith("InArgument("):
            direction = "In"
        elif row[1].startswith("OutArgument("):
            direction = "Out"
        type = row[1].split('(')[1][:-1]
        new_sheet.range(f"A{target_row}").value = row[0]  # DisplayName
        new_sheet.range(f"M{target_row}").value = direction  # Direction
        new_sheet.range(f"T{target_row}").value = type  # Type
        new_sheet.range(f"AR{target_row}").value = ARGS[f"{sheetKey}:{row[0]}"]  # Description

        print(f"add args {i+1} / {len(rows)} : {row[0]} ({direction})")

    countOfadd = len(rows) -1
    if countOfadd < 0:
        countOfadd = 0

    return countOfadd

def flow_len(text):
    numOfNL = text.count("<NL/>")
    return len(text) + numOfNL * 20  # <NL/>を改行に置き換えるので、2文字分の長さを追加  

def add_flow(app, rows, new_sheet, offset):
    start_row = 13+offset
    for i, row in enumerate(rows):
        insert_row = target_row = start_row + i
        
        if i > 0:
            new_sheet.range(f"A{insert_row}:BT{insert_row}").select()
            app.selection.insert(shift='down')  # 行を挿入
        else:
            new_sheet.range(f"M3").value = f"{row[0]}処理"
            new_sheet.range(f"A{start_row}:BT{start_row}").row_height = 41.7  # 行の高さを設定

        new_sheet.range(f"A{start_row}:BT{start_row}").select()
        app.api.Selection.Copy()  # 行をコピー
        new_sheet.range(f"A{target_row}:BT{target_row}").select()
        app.selection.paste # xlPasteAll

        height = new_sheet.range(f"A{start_row}:BT{start_row}").row_height
        new_sheet.range(f"A{target_row+1}:BT{target_row+1}").row_height = height  # 行の高さをコピー元と同じにする
        
        try:

            properties = row[3]
            if flow_len(properties) > 130:  # Propertiesの長さで行の高さを調整
                extended_height = (height / 2) * flow_len(properties) // 60
                #print (f"長いプロパティ: {properties[:40]} len:{flow_len(properties)} height:{extended_height}")
                new_sheet.range(f"A{target_row}:BT{target_row}").row_height = extended_height
            properties = properties.replace("<NL/>", "\n")  # 改行TAGを改行に置換
            # データ書き込み
            new_sheet.range(f"A{target_row}").value = i+1  # No.
            new_sheet.range(f"C{target_row}").value = row[0]  # DisplayName
            new_sheet.range(f"I{target_row}").value = row[1]  # Path
            new_sheet.range(f"AC{target_row}").value = row[2]  # Type
            new_sheet.range(f"AK{target_row}").value = properties # Properties
        except Exception as e:
            print(f"Error writing flow data at row {row[0]}:[{row}]: {e}")
            continue

def set_title(file_name, sheet):
    name = os.path.splitext(os.path.basename(file_name))[0].split('.')[0]
    match = re.search(r'\((.*?)\)', name)
    if match:
        name = match.group(1)
    sheet.range(f"O13").value = f"{name}"

def rm_rn_sheets(wb,senario_name):
    igonore_sheets = OTHER_SHEETS[senario_name]

    for sheet in wb.sheets:
        sheet_name = sheet.name
        if sheet_name in igonore_sheets:
            try:
                sheet.name  = igonore_sheets[sheet_name]  # シート名を変更
                print(f"シナリオ:{senario_name} シート '{sheet_name}' の名前を '{igonore_sheets[sheet_name]}' に変更しました")
            except Exception as e:
                print(f"シナリオ:{senario_name} シート '{sheet_name}' の変更に失敗しました: {e}")
        elif "-" in sheet.name:
            sheet_name = sheet.name
            try:
                sheet.delete()  # シートを削除
                print(f"シナリオ:{senario_name} シート '{sheet_name}' を削除しました")
            except Exception as e:
                print(f"シナリオ:{senario_name} シート '{sheet_name}' の削除に失敗しました: {e}")

def create_sheet(app, wb, csv_file , template_sheet, sheet_name, senario_name):
    try: 
        print(f"シート '{sheet_name}' を作成中...")
        # シートコピー（templateシートの直後にコピー）
        template_sheet.api.Copy(After=template_sheet.api)

        # コピー後、名前が "template (数字)" のシートを探す
        pattern = re.compile(r"template \(\d+\)")
        copied_sheets = [s for s in wb.sheets if pattern.fullmatch(s.name)]
        if not copied_sheets:
            print("エラー: コピー後の 'template (数字)' シートが見つかりません")
            return
        new_sheet = copied_sheets[0]
        new_sheet.name = sheet_name
        xaml = XAMLCsvReader(csv_file)
        srcdata = xaml.read()
         
        rm_rn_sheets(wb, senario_name)  # 不要なシートを削除

        offset = add_args(app, srcdata["args"], new_sheet, sheet_name)
        print(f"引数の追加完了: {offset} 行追加")
        add_flow(app, srcdata["flow"], new_sheet, offset)
        if sheet_name not in DESCRIPTIONS:
            sheet_name = f"{senario_name}:{sheet_name}"
        
        new_sheet.range("A7").value = DESCRIPTIONS[sheet_name]  # シートのタイトルを設定
         # シートの選択を解除
        print(f"シート '{sheet_name}' を作成しました")
        app.api.ActiveSheet.Select()
        app.api.ActiveSheet.Range("A1:A1").Select()

    except Exception as e:
        print(f"sheet name:{sheet_name} Error:{e}")
        raise e
    # finally:
    #     # シートの選択を解除
    #     print(f"シート '{sheet_name}' を作成しました")
    #     app.api.ActiveSheet.Select()
    #     app.api.ActiveSheet.Range("A1").Select()

def wrap_japanese_for_excel(text, max_chars):
    tagger = MeCab.Tagger()
    words = tagger.parse(text).split()
    lines, line = [], ""
    for word in words:
        if len(line + word) > max_chars:
            lines.append(line)
            line = word
        else:
            line += word
    lines.append(line)
    return "\n".join(lines)

def write_csv_to_excel(files, output_file, senario_name):

    app = xw.App(visible=False)
    try:
        wb = app.books.open(output_file)

        if "template" not in [s.name for s in wb.sheets]:
            print("エラー: 'template' シートが見つかりません")
            return

        set_title(output_file, wb.sheets["表紙"])
        for csv_file in files:
            if not os.path.exists(csv_file):
                print(f"エラー: CSVファイルが見つかりません: {csv_file}")
                continue
            template_sheet = wb.sheets["template"]
    
            # シート名をCSVファイル名から取得
            sheet_name = os.path.splitext(os.path.basename(csv_file))[0].split('-')[1]
            create_sheet(app, wb, csv_file, template_sheet, sheet_name, senario_name)

        #wb.sheets["表紙"].range("A1").select() # 表紙シートのA1セルを選択
        wb.sheets["template"].delete()  # templateシートを削除
        wb.save(output_file)
        print(f"Excelファイルを出力しました: {output_file}")

    finally:
        wb.close()
        app.quit()


def make_document(template_file, output_file, srcfiles):
    shutil.copy(template_file, output_file)
    print(f"Now generating document: {output_file} ....")
    matches = re.findall(r"\((.*?)\)", output_file)

    write_csv_to_excel(srcfiles, output_file,matches[0] if matches else "Unknown")


if __name__ == "__main__":
    #if len(sys.argv) != 4:
    #    print("使用法: python write_csv_to_template_xlwings.py 雛形.xlsx データ.csv 出力.xlsx")
    #    sys.exit(1)

    #write_csv_to_excel(sys.argv[1], sys.argv[2], sys.argv[3])
    template = "../template/doc_template.xlsx"
    docs = {"../dest/DD/DD01-RPA内部設計(AcMonthly04)_v0.8.0.xlsx": [
                "../dest/AcMonthly04-LoadKanjoJournal.csv",
                "../dest/AcMonthly04-GetAkikuraData.csv",
                "../dest/AcMonthly04-GetSettings.csv",
                "../dest/AcMonthly04-Main.csv"
                ],
            "../dest/DD/DD02-RPA内部設計(AcMonthly04B)_v0.8.0.xlsx": [
                "../dest/AcMonthly04B-LoadKanjoJournal.csv",
                "../dest/AcMonthly04B-GetAkikuraData.csv",
                "../dest/AcMonthly04B-GetSettings.csv",
                "../dest/AcMonthly04B-Main.csv"
                ],
            "../dest/DD/DD03-RPA内部設計(AcMonthly05)_v0.8.0.xlsx": [
                "../dest/AcMonthly05-LoadKanjoJournal.csv",
                "../dest/AcMonthly05-GetAkikuraData.csv",
                "../dest/AcMonthly05-GetSettings.csv",
                "../dest/AcMonthly05-Main.csv"
                ],
            "../dest/DD/DD04-RPA内部設計(AcMonthly05B)_v0.8.0.xlsx": [
                "../dest/AcMonthly05B-LoadKanjoJournal.csv",
                "../dest/AcMonthly05B-GetAkikuraData.csv",
                "../dest/AcMonthly05B-GetSettings.csv",
                "../dest/AcMonthly05B-Main.csv"
                ],
            "../dest/DD/DD05-RPA内部設計(AcMonthlyDataOutput)_v0.8.0.xlsx": [
                "../dest/AcMonthlyDataOutput-GetSettings.csv",
                "../dest/AcMonthlyDataOutput-Main.csv"
                ]
            }

    # docs = {"./dest/DD/AcMonthlyDataOutput.xlsx": [
    #             "./dest/AcMonthlyDataOutput-GetSettings.csv",
    #             "./dest/AcMonthlyDataOutput-Main.csv"
    #             ]
    #         }


    #make_document(template, "./dest/AcMonthly04.xlsx", docs["./dest/AcMonthly04.xlsx"])
    for k, v in docs.items():
        make_document(template, k, v)
    print("All documents created successfully.")
