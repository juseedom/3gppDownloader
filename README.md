# 3GPPDownloader
Download 3GPP pdf files from http://www.etsi.org/deliver/etsi_ts/xxx

Download 3GPP zip(doc) files from http://www.3gpp.org/ftp//Specs/archive/xxx

```shell
usage: 3GPPdownloader.py [-h] [-a | -d | -c] [-f {doc,pdf}] [-m MULTITHREAD]
                         [-r REL [REL ...]] [-s SERIES [SERIES ...]] [-p PATH]

3GPP download tools

optional arguments:
  -h, --help            show this help message and exit
  -a, --all             Perform Download, Convert and Merge if doc file type
  -d, --download        Perform Download only
  -c, --convertMerge    Perform Convert doc and Merge pdf file
  -f {doc,pdf}, --filetype {doc,pdf}
                        Downlaod 3GPP from website
  -m MULTITHREAD, --multithread MULTITHREAD
                        Multi-thread used, the default is 6
  -r REL [REL ...], --rel REL [REL ...]
                        Indicate the 3GPP release number, default is 15 (LTE:
                        8+, NR: 15+)
  -s SERIES [SERIES ...], --series SERIES [SERIES ...]
                        Download the 3GPP series number, eg. 36 for EUTRAN, 38
                        for NR
  -p PATH, --path PATH  Saving path for downloaded 3GPP
```

# Download doc from 3GPP

## Specific output folder

```shell
# spec download and save as .\36Series_Rel15\xxx.doc
python 3GPPdownloader.py -a -r 15 -s 36

# spec download and save as /user/xx/Desktop/36Series_Rel15/xxx.doc
python 3GPPdownloader.py -a -r 15 -s 36 -p /user/xx/Desktop/
```

## Download specific release and series

The series(-s) and release(-r) could accept multiple numbers
```shell
# download 36 series, latest version of relase 15
python 3GPPdownloader.py -a -r 15 -s 36

# download both 36 series release 15 and 38 series relase 15
python 3GPPdownloader.py -a -r 15 -s 36 38

# download 36 series, both release 14 and 15
python 3GPPdownloader.py -a -r 14 15 -s 36 38
```

## Download doc w/ or w/o convert to pdf

The sejda-console is used here:
https://github.com/torakiki/sejda/releases

```shell
# download 36 series, latest version of relase 15
# convert doc to pdf(only works under windows with Word installed)
python 3GPPdownloader.py -a -r 15 -s 36

# only perform convert doc to pdf
python 3GPPdownloader.py -c -r 15 -s 36

# only downlad doc
python 3GPPdownloader.py -d -r 15 -s 36
```
