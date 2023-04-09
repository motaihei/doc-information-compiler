import os
import subprocess
import datetime

# addores_output.ps1を実行して、開いているExcelファイルのアドレスを取得する
cmd = 'powershell.exe -ExecutionPolicy RemoteSigned -File addres_output.ps1'
result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)

# 取得したアドレスをマークダウン形式でリンクを持ったテキストにする
# PowerShellの出力はshift-jisなので、decodeで文字コードを変換する
addresses = result.stdout.decode('shift-jis').split()
markdown_text = ''
for address in addresses:
    file_name = address.split('\\')[-1]
    markdown_text += f"[{file_name}](file:///{address})\n"

# 'yyyy-mm-dd hh:mm:ss'形式で日付を初期化する
date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# 'yyyy-mm-dd'形式で日付を初期化する
date_file = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M')

# 出力先のディレクトリ
output_folder = '../output'

# 出力先
output = f'{output_folder}/file_addresses_{date_file}.md'

# マークダウン形式でリンクを持ったテキストをファイルに書き込む
with open(output, 'w', encoding='shift-jis') as file:
    # 日付を見出しにする
    file.write(f'# {date}\n')

    file.write(markdown_text)

    # fileの出力先の絶対パスを標準出力する
    print(os.path.abspath(file.name))







