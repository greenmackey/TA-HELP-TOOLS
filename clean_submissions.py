


import os
import glob



# フィードバックファイルを格納しているディレクトリのパス
print("フィードバックファイルを格納しているディレクトリの絶対パス:")
feedback_directory_path = input()

print('ファイルが変な文字を含んでいた場合，自動的に"-"に変更しますか？ YES=>y NO=>n: ', end="")
auto_correct = input()

files = glob.glob(feedback_directory_path + '/' + '[0-9]' * 2 + '-' + '[0-9]' * 6 + '*')

# PDF以外の拡張子がないかチェック
notpdf = set()
for file in files:
  file_name = os.path.basename(file)
  student_id = file_name[0:9]
  if file_name[-3:] != "pdf":
    notpdf.add(student_id)

# 学生証番号に重複がないか確認
multiple = set()
for id in range(len(files) - 1):
    # 学生証番号を取得（ファイル名の最初の9文字）
    file_name = os.path.basename(files[id])
    student_id = file_name[0:9]
    for id2 in range(id + 1, len(files)):
        file_name2 = os.path.basename(files[id2])
        student_id2 = file_name2[0:9]
        if student_id == student_id2:
            multiple.add(student_id)

# ファイル名に変な記号を含んでいないかチェック
include_invalid_string = set()
black_list = '¥ / : * ? " < > | % # ` { } ^ [ ]'.split()
for file in files:
    file_name = os.path.basename(file)
    new_file_name = file_name
    student_id = file_name[0:9]
    for black_letter in black_list:
        if black_letter in file_name:
          include_invalid_string.add(student_id)
          if auto_correct == 'y':
            new_file_name = new_file_name.replace(black_letter, "-")
    os.rename(file, os.path.join(feedback_directory_path, new_file_name))

# ファイル名を更新
files = glob.glob(feedback_directory_path + '/' + '[0-9]' * 2 + '-' + '[0-9]' * 6 + '*')

# ファイル名をエラーに応じて変更
for file in files:
  file_name = os.path.basename(file)
  student_id = file_name[0:9]
  error_message = ''
  if student_id in notpdf and 'notpdf' not in file_name:
    error_message += '-notpdf'
  if student_id in include_invalid_string and 'invalid_string' not in file_name:
    error_message += '-invalid_string'
  if student_id in multiple and 'multiple' not in file_name:
    error_message += '-multiple'
  os.rename(file, os.path.join(feedback_directory_path, student_id + error_message + file_name[9:]))

# エラーメッセージを表示
if len(notpdf) != 0:
  print("以下の学生にはフォーマットがPDFでないファイルがあります．")
  for student_id in notpdf:
    print(student_id)
if len(include_invalid_string) != 0:
  if auto_correct == 'y':
    print("以下の学生にはおかしな文字を含むファイルがありました．")
  else:
    print("以下の学生にはおかしな文字を含むファイルがあります．")
  for student_id in include_invalid_string:
    print(student_id)
if len(multiple) != 0:
  print("以下の学生には複数のファイルがあります．")
  for student_id in multiple:
    print(student_id)