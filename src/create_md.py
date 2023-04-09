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

# 'yyyy-mm-dd'形式で日付を文字列で初期化する
date = datetime.date.today().strftime('%Y-%m-%d')

# マークダウン形式でリンクを持ったテキストをファイルに書き込む
with open(f'../output/file_addresses_{date}.md', 'w', encoding='shift-jis') as file:
    # 見出しは日付にする
    file.write(f'# {date}\n')
    file.write(markdown_text)

# 出力先を標準出力にする
print(markdown_text)






