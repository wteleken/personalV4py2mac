from setuptools import setup

setup(
    app=['gerar_arquivo_novo_treino.py', 'criar_treino.py', 'unir_planilhas.py'],
    options={'py2app': {'packages': ['openpyxl', 'pandas', 'numpy', 'random']}},
    setup_requires=['py2app'],
)
