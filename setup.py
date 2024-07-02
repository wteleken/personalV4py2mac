from setuptools import setup

APP = ['gerar_arquivo_novo_treino.py']
DATA_FILES = []

OPTIONS = {
    'argv_emulation': True,
    'packages': ['os', 'openpyxl'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
