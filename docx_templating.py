# -*- coding: utf-8 -*-
'''
Created : 2020-03-22

@author: Andrey Luzhin
'''
import re
from docx import Document
from lxml import etree

HEADER_FOOTER = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")

def delete_runs_tags(text):
    """ Delete Word <w:t> tags (and everything between them) inside patterns before replace """
    return re.sub(r'(?:\{(?:<[^>\{\}]*>)*\{)([^\{\}]+?)(?:\}(?:<[^>\{\}]*>)*\})',
                  lambda x : re.sub('<\/w:t>[^\{\}]*?(?:<w:t>|<w:t [^>]*>)','', x.group(), flags=re.DOTALL),
                  text, flags=re.DOTALL)

def escape(str):
    """ Escape special characters """
    str = str.replace("&", "&amp;") # Should be the first. In order not to break next replacements.
    str = str.replace("<", "&lt;")
    str = str.replace(">", "&gt;")
    str = str.replace("'", "&apos;")
    str = str.replace("\"", "&quot;")
    return str

def replace_in_xml(xml, replace_dict):
    """ Replace objects in the document """
    text = etree.tostring(xml, encoding='unicode', pretty_print=False)
    text = delete_runs_tags(text)
    #print(re.findall(r'\{\{((?:[^\{\}])+?)\}\}', text, flags=re.DOTALL))
    text = re.sub(r'\{\{((?:[^\{\}])+?)\}\}',
                  lambda x: escape(replace_dict.get(x.group(1),
                                                    'UNKNOWN PATTERN "'+x.group(1)+'"')),
                  text, flags=re.DOTALL)
    return etree.fromstring(text)

substitutes = {
    'hello' : 'Hello! Привет! ¡Hola!',
    'test'  : 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
    'bye'   : 'Bye! Пока! ¡Adiós!',
    'chars' : '[ < > ? . * { } & % " \' ]',
}

tpl = Document('template.docx')

body = tpl.element.body
rels = tpl.part.rels

# Replacing objects in headers & footers
for relKey, val in rels.items():
    if val.reltype in HEADER_FOOTER:
        rels[relKey]._target._element=replace_in_xml(val._target._element, substitutes)

# Replacing in document body
tpl.element.replace(body,replace_in_xml(body, substitutes))

tpl.save('result.docx')
