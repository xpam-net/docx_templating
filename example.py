# -*- coding: utf-8 -*-
from docxt import DocxT

substitutes = {
    'hello': 'Hello! Привет! ¡Hola!',
    'test': 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
    'bye': 'Bye! Пока! ¡Adiós!',
    'chars': '[ < > ? . * { } & % " \' ]',
    'blank': ''
    # used to demonstate how to preserve some text inside double curly braces. Just insert {{blank}} somewhere inside such text
}

# Opening template file
tpl = DocxT('example_template.docx')
# show all patterns inside the document
for i, t in enumerate(tpl.get_all_tags()):
    print(i, ': ' + t)
# now replacing
tpl.replace_all(substitutes)
# and now saving as new file
tpl.save_file('example_result.docx')
