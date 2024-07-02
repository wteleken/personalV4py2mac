from setuptools import setup

APP = ['gerar_arquivo_novo_treino.py', 'criar_treino.py', 'Unir_planilhas.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['openpyxl', 'pandas', 'random', 'numpy', 'os', 'datetime'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)