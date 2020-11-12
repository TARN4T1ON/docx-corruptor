#!/usr/bin/env python3

import os
from shutil import rmtree
from distutils.dir_util import copy_tree

from configparser import ConfigParser as configparser

from xml.etree import ElementTree as elementtree

import random
import datetime

import re

from zipfile import ZipFile as zipfile

root = os.path.dirname(os.path.realpath(__file__))

settings = {}

config = configparser()
config.read(os.path.join(root, "settings.cfg"))

configValues = [
        "DocumentName",

        "Application",
        "AppVersion",
        "Revision",
        "Template",
        "DocSecurity",

        "TotalTime",

        "Pages",
        "Words",
        "Characters",
        "CharactersWithSpaces",
        "Lines",
        "Paragraphs",

        "Company",
        "Creator",
        "LastModifiedBy",

        "Title",
        "Subject",
        "Keywords",
        "Description",

        "Created",
        "Modified",

        "ScaleCrop",
        "LinksUpToDate",
        "SharedDoc",
        "HyperlinksChanged",
]

for configValue in configValues:
    settings[configValue] = config["CORRUPTOR"][configValue]

documentDirectory = os.path.join(root, settings["DocumentName"])

documentPropertyPaths = {
    "app": ["docProps", "app.xml"],
    "core": ["docProps", "core.xml"],

    "document": ["word", "document.xml"],
}

xmlNamespaces = {
    "app": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",

    "core": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",

    "document": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

for namespace in xmlNamespaces:
    elementtree.register_namespace(namespace, xmlNamespaces[namespace])

def main():
    rmtree(documentDirectory, ignore_errors=True)

    copy_tree("./doc-template", documentDirectory)

    # write core xml from settings

    corePath = os.path.join(documentDirectory, *documentPropertyPaths["core"])
    core = elementtree.parse(corePath)

    coreRoot = core.getroot()
    for coreEl in coreRoot.iter():
        coreElTag: str = coreEl.tag
        _coreElTag = coreElTag.lower()

        for setting in settings:
            _setting: str = setting.lower()

            if _coreElTag.endswith(_setting):
                coreEl.text = str(settings[setting])

    core.write(corePath, xml_declaration=True)

    # write app xml from settings

    appPath = os.path.join(documentDirectory, *documentPropertyPaths["app"])
    app = elementtree.parse(appPath)

    appRoot = app.getroot()
    for appEl in appRoot.iter():
        appElTag: str = appEl.tag
        _appElTag = appElTag.lower()

        for setting in settings:
            _setting: str = setting.lower()

            if _appElTag.endswith(_setting):
                appEl.text = str(settings[setting])

    app.write(appPath, xml_declaration=True)

    # write doc xml

    docPath = os.path.join(documentDirectory, *documentPropertyPaths["document"])
    doc = elementtree.parse(docPath)

    docRoot = doc.getroot()

    docBody = docRoot.find("document:body", xmlNamespaces)

    docText = ""
    for i in range(0, config.getint("CORRUPTOR", "FileSize") * 1000):
        docText += chr(random.randint(32, 127))

    docBody.text = docText

    doc.write(docPath, xml_declaration=True)

    #corrupt doc

        # read doc to docData
    _docRead = open(docPath, "r", encoding = "ascii")

    docText = _docRead.read()
    docData = bytearray(docText, encoding="ascii")

    _docRead.close()

        # get header size
    docHeader = re.match(r".*?<[^/\s]*?document.*?>", docText, re.DOTALL | re.MULTILINE)

        # corrupt header
    for i in range(0, docHeader.end()):
        if i > len(docData):
            break

        docData[i] = random.randint(0, 255)

        # randomly corrupt more of doc
    for i in range(0, 1000):
        index = 1000 + int(random() * (len(docData) - 1))

        if index > len(docData) - 1:
            continue

        docData[index] = random.randint(0, 255)

    # write corrupt doc

    docCorrupt = open(docPath, "wb")

    docCorrupt.write(docData)

    docCorrupt.close()

    # zip file into .docx to be sent :)

    outPath = os.path.join(documentDirectory, settings["DocumentName"] + ".docx")
    outFile = zipfile(outPath, "w")

    for _directory, _directories, _files in os.walk(documentDirectory):
        _directoryRelative = os.path.relpath(_directory, documentDirectory)

        for _file in _files:
            if _file == os.path.basename(outPath):
                continue

            _filePath = os.path.join(_directory, _file)
            outFile.write(_filePath, os.path.join(_directoryRelative, _file))

    outFile.close()

    return


main()
