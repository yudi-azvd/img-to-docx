# -*- coding: utf-8 -*-
# https://python-docx.readthedocs.io/en/latest/

# Configurar layout da página
# https://python-docx.readthedocs.io/en/latest/user/sections.html

# Pesquisar: 
# 1 - allow docx file to be modified while open
# 2 - Word night mode
import os

FOLDER = 'imgs'

images = os.listdir(FOLDER)

image = images[0]

print(image)
