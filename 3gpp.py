#!/usr/bin/python3

"""
automatically download 3gpp specs
"""

import logging
import re
import shutil
import sys

from functools import partial
from multiprocessing import Pool
from pathlib import Path
from typing import Iterator, List, Text, Union
from urllib import request
from zipfile import ZipFile

__all__ = ['HEADERS', 'NUMS', 'download']

HEADERS = {
    'User-Agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36'
}

NUMS = 8


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
    print('parsing', url, 'now...')
    req = request.Request(url, headers=HEADERS)
    response = str(request.urlopen(req).read())
    # example of response:
    # <A HREF="/deliver/etsi_ts/138200_138299/138202/">[To Parent Directory]</A>
    for link in re.compile(r'<A HREF="([^"]+)">[^\[^<]+</A>', re.IGNORECASE).findall(response):
        if link.endswith('/'):
            yield (url + '/' + link.split('/')[-2])
        else:
            yield (url + '/' + link.split('/')[-1])


def multi_processing(spec_url: Text, release: Text, dir_path: Union[Path, Text] = '.', extract_zip: bool = True) -> None:
    """ be a multiprocess wrapper to download specs

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


def download(release: int = 15, series: int = 38, dir_path: Union[Path, Text] = '.', extract_zip: bool = True) -> None:
    """ download the 3gpp sepcs, according to release number and series number

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
    _func = partial(multi_processing, release=release, dir_path=dir_path, extract_zip=extract_zip)
    with Pool(NUMS) as pool:
        for _ in pool.imap_unordered(_func, urls):
            pass


def convert_pdf(dir_path: Union[Path, Text] = '.'):
    pass

if __name__ == '__main__':
    from argparse import ArgumentParser
    parser = ArgumentParser('3gpp download')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('-a', '--all', action='store_true', help='Perform Download, Extract and Convert')
    mode.add_argument('-d', '--download', action='store_true', help='Perform Download only')
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

    for release in args.release:
        for series in args.series:
            dir_path = Path(args.path)/'{}Series_Rel{}'.format(series, release)
            if not Path(dir_path).exists():
                Path(dir_path).mkdir()

            if args.all:
                download(int(release), int(series), dir_path, True)
            elif args.download:
                download(int(release), int(series), dir_path, False)
