from setuptools import setup

APP = ['gerar_arquivo_novo_treino.py', 'criar_treino.py', 'unir_planilhas.py']
OPTIONS = {
    'packages': ['openpyxl', 'pandas', 'numpy'],
    'includes': [],
    'excludes': [],
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
