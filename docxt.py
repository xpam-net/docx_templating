# -*- coding: utf-8 -*-
'''
Class DocxT

Created : 2020-03-24

@author: Andrey Luzhin
'''
import re
from docx import Document
from lxml import etree

class DocxT:
    """ Class for templates editing """
    # headers and footers reltypes
    HEADER_FOOTER = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")

    def __init__(self, filename):
        """ constructor """
        self.filename = filename
        self.file = Document(filename)
        self.body = self.file.element.body
        self.rels = self.file.part.rels
        self.regx = re.compile(r'\{\{((?:[^\{\}])+?)\}\}', flags=re.DOTALL)

    def __retr__(self):
        """ string representation of an object """
        return '{DocxT: "' + self.filename + '"}'

    def get_docx(self):
        """ link to docx.Document object """
        return self.file

    def get_template_name(self):
        """ filename of template """
        return self.filename

    def get_headers_footers(self):
        """ get rels for headers and footers """
        for relKey, val in self.rels.items():
            if val.reltype in self.HEADER_FOOTER:
                yield relKey, val

    def save_file(self, filename):
        """ save resulting file """
        self.file.save(filename)

    def escape(self, str):
        """ Escape special characters """
        str = str.replace("&", "&amp;") # Should be the first. In order not to break next replacements.
        str = str.replace("<", "&lt;")
        str = str.replace(">", "&gt;")
        str = str.replace("'", "&apos;")
        str = str.replace('"', "&quot;")
        return str

    def xml_to_string(self, xml):
        """ xml object to string """
        return etree.tostring(xml, encoding='unicode', pretty_print=False)

    def string_to_xml(self, str):
        """ string object to xml """
        return etree.fromstring(str)

    def delete_runs_tags(self, text):
        """ Delete Word <w:t> tags (and everything between them) inside patterns before replace """
        return re.sub(r'(?:\{(?:<[^>\{\}]*>)*\{)([^\{\}]+?)(?:\}(?:<[^>\{\}]*>)*\})',
                      lambda x : re.sub('<\/w:t>[^\{\}]*?(?:<w:t>|<w:t [^>]*>)','', x.group(), flags=re.DOTALL),
                      text, flags=re.DOTALL)

    def replace_in_xml(self, xml, replace_dict):
        """ Replace patterns in xml """
        text = self.xml_to_string(xml)
        text = self.delete_runs_tags(text)
        text = self.regx.sub(lambda x: self.escape(replace_dict.get(x.group(1), 'UNKNOWN PATTERN "'+x.group(1)+'"')), text)
        return self.string_to_xml(text)

    def replace_in_headers(self, replace_dict):
        """ Replace inside haders and footers """
        for relKey, val in self.get_headers_footers():
            # Not sure if editing headers is now safe with the python-docx module. Better apply templating before final saving.
            self.rels[relKey]._target._element=self.replace_in_xml(val._target._element, replace_dict)

    def replace_in_body(self, replace_dict):
        """ Replace inside inside the document body """
        self.file.element.replace(self.body, self.replace_in_xml(self.body, replace_dict))

    def replace_all(self, replace_dict):
        """ Replace inside the entire document """
        self.replace_in_headers(replace_dict)
        self.replace_in_body(replace_dict)

    def get_body_tags(self):
        """ Set of tags in the document body """
        return set(self.regx.findall(self.delete_runs_tags(self.xml_to_string(self.body))))

    def get_header_footer_tags(self):
        """ Set of tags inside headers and footers """
        s = set()
        for _, val in self.get_headers_footers():
            s = s.union(set(self.regx.findall(self.delete_runs_tags(self.xml_to_string(val._target._element)))))
        return s

    def get_all_tags(self):
        """ Set of all tags """
        return self.get_body_tags().union(self.get_header_footer_tags())
