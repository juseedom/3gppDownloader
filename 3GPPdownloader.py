#!/usr/bin/python3

"""
automatically download 3gpp specs
"""

import logging
import os
import re
import shutil
import sys
from functools import partial
from multiprocessing import Pool
from pathlib import Path
from typing import Iterator, List, Text, Union
from urllib import request
from zipfile import ZipFile

__all__ = ['HEADERS', 'NUMS', 'm_download', 'm_convert_pdf']

HEADERS = {
    'User-Agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36'
}

NUMS = 8

SEJDA_PATH = str((Path(__file__).parent / 'bin/sejda-console.bat').absolute())

global PLATFORM

if sys.platform.startswith('win32'):
    PLATFORM = 'win'
    import win32com.client as win32
elif sys.platform.startswith('darwin'):
    PLATFORM = 'mac'
elif sys.platform.startswith('linux'):
    PLATFORM = 'linux'
else:
    PLATFORM = ''


def url_download(url: Text, dir_path: Union[Path, Text] = '.', extract_zip: bool = True) -> None:
    """ download the url and save as files

    Args:
        url (str): the file url to be downloaded
        dir_path (Path, str): folder location to be saved
        extract_zip (bool): extract zip files, Default is True, extract
    """
    file_name = Path(dir_path) / url.split('/')[-1]
    if file_name.exists():
        print('Found file exists, overwriting: %s' % str(file_name))
    
    with request.urlopen(url) as response, open(file_name, 'wb') as out_file:
        shutil.copyfileobj(response, out_file)
        print('File Downloaded: %s' % str(file_name))

    if file_name.suffix == '.zip' and extract_zip:
        ZipFile(file_name).extractall(dir_path)
        file_name.unlink()
        print(file_name.name, 'extracted...')


def url_load(url: Text) -> Iterator:
    """ load a url location and return links

    Args:
        url (str): the url location to be read

    """
    # print('parsing', url, 'now...')
    req = request.Request(url, headers=HEADERS)
    response = str(request.urlopen(req).read())
    # example of response:
    # <A HREF="/deliver/etsi_ts/138200_138299/138202/">[To Parent Directory]</A>
    for link in re.compile(r'<A HREF="([^"]+)">[^\[^<]+</A>', re.IGNORECASE).findall(response):
        if link.endswith('/'):
            yield (url + '/' + link.split('/')[-2])
        else:
            yield (url + '/' + link.split('/')[-1])


def download(spec_url: Text, release: Text, dir_path: Union[Path, Text] = '.', extract_zip: bool = True) -> None:
    """ download the 3gpp sepcs, according to release number and series number

    Args:
        spec_url (str): the spec url to be download, e.g. https://www.3gpp.org/ftp/Specs/archive/38_series/38.900
        release (str): release version of spec to be download
        dir_path (Path, str): download save path for spec, Default is '.' current directory
        extract_zip (bool): extract zip files, Default is True, extract
    """
    rel_url = [x for x in url_load(spec_url) if x.split('-')[-1].startswith(release)]
    if rel_url:
        rel_url.sort()
        # download the latest spec within specific release
        rel_url = rel_url[-1]
        if rel_url.split('.')[-1] == 'zip':
            url_download(rel_url, dir_path, extract_zip)


def m_download(release: int = 15, series: int = 38, dir_path: Union[Path, Text] = '.', extract_zip: bool = True) -> None:
    """ multiprocess wrapper to download specs

    Args:
        release (int): the release number to be download, Default is 15, first release of NR
        series (int): the series number to be download, Default is 38, NR release
        dir_path (Path, str): download save path for spec, Default is '.' current directory
        extract_zip (bool): extract zip files, Default is True, extract
    """

    series = str(series)
    release = '0123456789abcdefghijklmn'[release]
    start_url = 'https://www.3gpp.org/ftp/Specs/archive/{}_series'
    print('Loading specs urls...')
    urls = []
    for url in url_load(start_url.format(series)):
        if url.split('/')[-1].startswith(series):
            urls.append(url)
    print('Specs urls loaded...')
    _func = partial(download, release=release, dir_path=dir_path, extract_zip=extract_zip)
    with Pool(NUMS) as pool:
        for _ in pool.imap_unordered(_func, urls):
            pass


def convert_pdf(doc_path: Union[Path, Text]) -> None:
    """ convert the word *.doc, *.docx to *.pdf (save to same location)

    Args:
        doc_path (str): the file path for doc files
    """
    try:
        word = None
        doc = None
        if Path(doc_path).suffix in ['.doc', '.docx']:
            output = Path(doc_path).with_suffix('.pdf')
            if output.exists():
                print('PDF file existed: {}'.format(doc_path))
                return
            print('Convert WORD into PDF: ', str(doc_path))
            word = win32.DispatchEx('Word.Application')
            word.Visible = 0
            # https://docs.microsoft.com/en-us/office/vba/api/word.documents.open
            # .Open (FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, 
            # PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, 
            # WritePasswordTemplate, Format, Encoding, Visible, OpenConflictDocument, 
            # OpenAndRepair, DocumentDirection, NoEncodingDialog)
            doc = word.Documents.Open(str(doc_path), False, False, True)
            # 'OutputFileName', 'ExportFormat', 'OpenAfterExport', 'OptimizeFor', 'Range',
            # 'From', 'To', 'Item', 'IncludeDocProps', 'KeepIRM', 'CreateBookmarks', 'DocStructureTags',
            # 'BitmapMissingFonts', 'UseISO19005_1', 'FixedFormatExtClassPtr'
            doc.ExportAsFixedFormat(
                str(output),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=1)
    except Exception as e:
        print('Open failed due to ' + str(e))
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()


def m_convert_pdf(dir_path: Union[Path, Text] = '.') -> None:
    """ multiprocessing wrapper for convert_pdf

    Args:
        dir_path (Path, str): download save path for spec, Default is '.' current directory
    """
    print('Convert doc to pdf under', dir_path)
    files = []
    for _file in Path(dir_path).glob('*.doc*'):
        if not _file.with_suffix('.pdf').exists():
            if _file.name.startswith('~'):
                continue
            files.append(_file.absolute())

    with Pool(NUMS) as pool:
        for _ in pool.imap_unordered(convert_pdf, files): ...
    
    if not Path(SEJDA_PATH).exists():
        print('Cannot find sejda at ', SEJDA_PATH)
        merge_pdf2(dir_path)
    else:
        merge_pdf(dir_path)


def merge_pdf2(dir_path: Union[Path, Text] = '.') -> None:
    """ merge multiple pdf files into one file

    Args:
        dir_path (Path, Text): input pdf files path, Default is '.'
    """
    from PyPDF2 import PdfFileMerger
    pdf_files = [x for x in Path(dir_path).glob('*.pdf')]
    _pdf_files = [x.name.split('_')[0] for x in pdf_files]
    merge_pdf = {x: [y for y in pdf_files if y.name.startswith(x+'_')] 
                    for x in _pdf_files if _pdf_files.count(x) > 1}
    # print(merge_pdf)
    for spec in merge_pdf:
        print('Start to merge %s.pdf with pypdf2' %spec)
        files = sorted(merge_pdf[spec])
        for _file in files:
            if 'cover' in _file.name:
                _cover = _file
                files.remove(_cover)
                break
        else:
            _cover = None
        if _cover:
            files.insert(0, _cover)

        out_file = str((Path(dir_path)/spec).with_suffix('.pdf'))
        # print(files, out_file)

        merger = PdfFileMerger()
        for pdf in files:
            merger.append(open(str(pdf), 'rb'))
        with open(out_file, 'wb') as _out_pdf:
            merger.write(_out_pdf)
            merger.close()


def merge_pdf(dir_path: Union[Path, Text] = '.', remove: bool = False) -> None:
    """ merge multiple pdf files into one file

    Args:
        dir_path (Path, Text): input pdf files path, Default is '.'
        remove (bool): remove original pdf files after conversion, Default is False
    """
    from subprocess import call
    pdf_files = [x for x in Path(dir_path).glob('*.pdf')]
    _pdf_files = [x.name.split('_')[0] for x in pdf_files]
    merge_pdf = {x: [y for y in pdf_files if y.name.startswith(x)] for x in _pdf_files if _pdf_files.count(x) > 1}
    # print(merge_pdf)
    for spec in merge_pdf:
        print('Start to merge %s.pdf with sejda' %spec)
        cmd = [SEJDA_PATH, 'merge', '--files']
        files = sorted(merge_pdf[spec])
        for _file in files:
            if 'cover' in _file.name:
                _cover = _file
                files.remove(_cover)
                break
        else:
            _cover = None
        if _cover:
            files.insert(0, _cover)
        out_file = str((Path(dir_path)/spec).with_suffix('.pdf'))
        [cmd.append(str(x)) for x in files]
        cmd.append('--output')
        cmd.append(out_file)
        cmd.append('--overwrite')
        # print(cmd)
        if (call(cmd, timeout=120) == 0) and remove:
            for x in merge_pdf[spec]:
                Path(x).unlink()


if __name__ == '__main__':
    from argparse import ArgumentParser
    parser = ArgumentParser('3gpp download')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('-a', '--all', action='store_true', help='Perform Download, Extract and Convert')
    mode.add_argument('-d', '--download', action='store_true', help='Perform Download only')
    mode.add_argument('-c', '--convert', action='store_true', help='Perform Convert word to pdf (win+word only)')
    parser.add_argument('-m', '--multithread', type=int, default=8,
                        help='Multi-thread used, the default is 8')
    parser.add_argument('-r', '--release', type=str, nargs='+', default=['15', ],
                        help='Indicate the 3GPP release number, default is 15 (LTE: 8+, NR: 15+)')
    parser.add_argument('-s', '--series', type=str, nargs='+', default=['38', ],
                        help='Download the 3GPP series number, eg. 36 for EUTRAN, 38 for NR')
    parser.add_argument('-p', '--path', type=str, default='.',
                        help='Saving path for downloaded 3GPP')

    args = parser.parse_args()
    NUMS = args.multithread
    _path = Path(args.path).absolute()

    for release in args.release:
        for series in args.series:
            dir_path = _path/'{}Series_Rel{}'.format(series, release)
            if not Path(dir_path).exists():
                Path(dir_path).mkdir()

            if args.all:
                m_download(int(release), int(series), dir_path, True)
            elif args.download:
                m_download(int(release), int(series), dir_path, False)

            if (PLATFORM == 'win') and (args.all or args.convert):
                m_convert_pdf(dir_path)
