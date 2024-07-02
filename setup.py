from setuptools import setup

# Para o script gerar_arquivo_novo_treino.py
setup(
    app=['gerar_arquivo_novo_treino.py'],
    options={'py2app': {'packages': ['openpyxl']}},
    setup_requires=['py2app'],
)

# Para o script criar_treino.py
setup(
    app=['criar_treino.py'],
    options={'py2app': {'packages': ['pandas', 'openpyxl', 'numpy', 'random']}},
    setup_requires=['py2app'],
)

# Para o script unir_planilhas.py
setup(
    app=['unir_planilhas.py'],
    options={'py2app': {'packages': ['openpyxl', 'pandas']}},
    setup_requires=['py2app'],
)
