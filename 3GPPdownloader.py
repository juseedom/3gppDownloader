# pdf version:
# http://www.etsi.org/deliver/etsi_ts/136100_136199/136101/12.07.00_60/ts_136101v120700p.pdf
#
# zip(doc) version:
# http://www.3gpp.org/ftp//Specs/archive/36_series/36.331/36331-e30.zip
import os
import re
from multiprocessing import Pool
from pathlib import Path
from urllib import request
from zipfile import ZipFile

import win32com.client as win32
NUM = 8

def urldownlaod(urlstr, save_path='.'):
    """ download the url and save as files
    
    Args:
        urlstr (str): the url need to be downloaded
        save_path (str): the save location for download urls

    """
    headers = {
        'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'
    }
    url = request.Request(url=urlstr, headers=headers)
    dl_file = request.urlopen(url)
    name = Path(save_path) / urlstr.split('/')[-1]
    with open(name, 'wb') as save_file:
        save_file.write(dl_file.read())
    print(name.name + ' downloaded...')
    if name.suffix == '.zip':
        ZipFile(name).extractall(path=save_path)
        name.unlink()
        print(name.name + ' extracted...')

def urlload(urlstr):
    """ load a url location and return the links observed
    
    Args:
        urlstr (str): the url location to read

    Returns:
        result (list): a list of link founded in this url
    """

    result = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'
        }
    url = request.Request(url=urlstr, headers=headers)
    response = str(request.urlopen(url).read())
    # <A HREF="/deliver/etsi_ts/138200_138299/138202/">[To Parent Directory]</A>
    # remove the link for [To Parent Directory]
    href = re.compile('<A HREF="([^"]+)">[^\[^<]+</A>')
    for link in href.findall(response):
        if link.endswith('/'):
            result.append(urlstr + '/' + link.split('/')[-2])
        else:
            result.append(urlstr + '/' + link.split('/')[-1])
    return result

def multi_processing(spec_url, rel, save_path='.', mode='doc', convert=True):
    if mode == 'doc':
        rel_url = [x for x in urlload(spec_url) if x.split('-')[-1].startswith(rel)]
        if rel_url:
            rel_url.sort()
            rel_url = rel_url[-1]
            # download zip only or all files
            if rel_url.split('.')[-1] == 'zip':
                # urldownlaod(rel_url)
                headers = {
                    'User-Agent':
                    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'
                }
                url = request.Request(url=rel_url, headers=headers)
                dl_file = request.urlopen(url)
                file_name = Path(save_path) / rel_url.split('/')[-1]
                with open(file_name, 'wb') as save_file:
                    save_file.write(dl_file.read())
                print(file_name.name + ' downloaded...')
                if file_name.suffix == '.zip':
                    ZipFile(file_name).extractall(path=save_path)
                    file_name.unlink()
                    print(file_name.name + ' extracted...')
                # if convert:
                #     for doc_file in Path(save_path).glob('{}*.doc*'.format(file_name.name[:5])):
                #         word2pdf(doc_file.absolute())

def word2pdf(filedoc):
    """ convert the word *.doc, *.docx to *.pdf (save to same location)

    Args:
        filedoc (str): the file path for doc files
    """
    try:
        word = None
        doc = None
        if Path(filedoc).suffix in ['.doc', '.docx']:
            output = Path(filedoc).with_suffix('.pdf')
            if output.exists():
                print('PDF file existed: {}'.format(filedoc))
                return
            print('Convert WORD into PDF: ', str(filedoc))
            word = win32.DispatchEx('Word.Application')
            word.Visible = 0
            doc = word.Documents.Open(str(filedoc), False, False, True)
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
        print('Open failed due to \n' + str(e))
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()

def download3GPP(file_type='doc', rel=13, series=36, path='.', convert=True):
    """ download the 3gpp spec, note the timer guard is 15min

    Args:
        file_type (str): different file type
        rel (int): the release number of sepc
        series (int): the series number of spec
        path (str): save location
    """
    # os.chdir(path)
    series = str(series)
    if file_type == 'pdf':
        # https://www.etsi.org/deliver/etsi_ts/138100_138199/13810101/15.02.00_60/
        rel = str(rel)
        str_url = 'http://www.etsi.org/deliver/etsi_ts'
        for str_url in [x for x in urlload(str_url) if series == x.split('/')[-1][1:3]]:
            for spec_url in urlload(str_url):
                # decide which rel to download
                rel_url = [x for x in urlload(spec_url) if rel == x.split('/')[-1][:len(rel)]]
                if rel_url:
                    rel_url.sort()
                    rel_url = rel_url[-1]
                    # download pdf only or all files
                    file_fmt = 'pdf'
                    for f_url in [a for a in urlload(rel_url) if a.split('.')[-1] == file_fmt]:
                        urldownlaod(f_url, path)
    elif file_type == 'doc':
        str_url = 'http://www.3gpp.org/ftp/Specs/archive/{}_series'
        release = '0123456789abcdefghijk'
        rel = int(rel)
        rel = release[rel]
        spec_urls = urlload(str_url.format(series))
        with Pool(NUM) as pool:
            # pool.starmap(multi_processing, [(x, rel, path, 'doc') for x in spec_urls], convert)
            processes = [
                pool.apply_async(multi_processing, arg)
                for arg in [(x, rel, path, 'doc', convert) for x in spec_urls]
            ]
            for res in processes:
                result = res.get(timeout=900)
                if result:
                    print(result)

def merge_specs(spec_path='.', sejda_path=r'..\bin\sejda-console.bat', remove=True):
    from subprocess import call
    pdf_files = [x for x in os.listdir(spec_path) if x.endswith('.pdf')]
    _pdf_files = [x.split('_')[0] for x in pdf_files]
    merge_pdf = {x: [y for y in pdf_files if y.startswith(x)] for x in _pdf_files if _pdf_files.count(x) > 1}
    # print(merge_pdf)
    old_wd = os.curdir
    os.chdir(spec_path)
    for spec in merge_pdf:
        print('Start to merge %s.pdf' %spec)
        cmd = [sejda_path, 'merge', '--files']
        [cmd.append(str(x)) for x in merge_pdf[spec]]
        cmd.append('--output')
        cmd.append('%s.pdf' %spec)
        if (call(cmd, timeout=120) == 0) and remove:
            for x in merge_pdf[spec]:
                Path(x).unlink()
    os.chdir(old_wd)

if __name__ == '__main__':
    from argparse import ArgumentParser
    parser = ArgumentParser(description='3GPP download tools')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('-a', '--all', action='store_true', help='Perform Download, Convert and Merge if doc file type')
    mode.add_argument('-d', '--download', action='store_true', help='Perform Download only')
    mode.add_argument('-c', '--convertMerge', action='store_true', help='Perform Convert doc and Merge pdf file')
    parser.add_argument(
        '-f', '--filetype',
        type=str,
        default='doc',
        choices=['doc', 'pdf'],
        help='Downlaod 3GPP from website')
    parser.add_argument(
        '-m', '--multithread',
        type=int,
        default=6,
        help='Multi-thread used, the default is 6')
    parser.add_argument(
        '-r', '--rel',
        type=str,
        nargs='+',
        default=['15', ],
        help='Indicate the 3GPP release number, default is 15 (LTE: 8+, NR: 15+)')
    parser.add_argument(
        '-s', '--series',
        type=str,
        nargs='+',
        default=['38', ],
        help='Download the 3GPP series number, eg. 36 for EUTRAN, 38 for NR')
    parser.add_argument(
        '-p', '--path',
        type=str,
        default='.',
        help='Saving path for downloaded 3GPP')

    args = parser.parse_args()
    file_type = args.filetype
    NUM=args.multithread

    for release in args.rel:
        for series in args.series:
            _path = Path(args.path)/'{}Series_Rel{}'.format(series, release)
            if not Path(_path).exists():
                Path(_path).mkdir()
            # print(file_type, path, convert)
            if args.all:
                # download doc word from 3gpp
                download3GPP(file_type, release, series, _path)
                while True:
                    files = [
                        str(x.absolute())
                        for x in Path(_path).glob(str(series) + '*.doc*')
                    ]
                    # remove existed pdf files
                    files = [
                        x for x in files
                        if not Path(x).with_suffix('.pdf').exists()
                    ]
                    if len(files) == 0:
                        break
                    with Pool(NUM) as pool:
                        # pool.map(word2pdf, files)
                        processes = [
                            pool.apply_async(word2pdf, (arg,)) for arg in files
                        ]
                        for res in processes:
                            result = res.get(timeout=900)
                            if result:
                                print("ERROR:")
                                print(result)

                # merge together
                # ../bin/sejda-console merge --files /Users/edi/Desktop/test.pdf /Users/edi/Desktop/test1.pdf --output /Users/edi/Desktop/merged.pdf
                merge_specs(_path)

            elif args.download:
                download3GPP(file_type, release, series, _path, False)

            elif args.convertMerge:
                while True:
                    files = [str(x.absolute()) for x in Path(_path).glob(str(series)+'*.doc*')]
                    # remove existed pdf files
                    files = [x for x in files if not Path(x).with_suffix('.pdf').exists()]
                    if len(files) == 0:
                        break
                    with Pool(NUM) as pool:
                        # pool.map(word2pdf, files)
                        processes = [
                            pool.apply_async(word2pdf, (arg, )) for arg in files
                        ]
                        for res in processes:
                            result = res.get(timeout=900)
                            if result:
                                print("ERROR:")
                                print(result)

                merge_specs(_path)
