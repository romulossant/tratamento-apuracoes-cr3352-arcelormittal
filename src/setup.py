# setup.py
import PyInstaller.__main__

PyInstaller.__main__.run([
    '--name=Tratamento Apurações IOS - CR 3352 ArcelorMittal',
    '--onefile',
    '--console',
    '--icon=icon.ico',
    'tratamento_apuracoes.py'
])
