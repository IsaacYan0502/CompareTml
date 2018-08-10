#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'compare module'
__author__ = 'JiaYou Yan'

import getopt
import logging
import os
import shutil
import sys
import time
import xml.dom.minidom
from functools import wraps

import lxml
import requests
import xlrd
import xlwt
from bs4 import BeautifulSoup
from progressbar import *
from xlrd import open_workbook
from xlutils.copy import copy
"""It was authored in May 2018"""


def function_timer(function):
    @wraps(function)
    def wrap_function(*args, **kwargs):
        t0 = time.time()
        result = function(*args, **kwargs)
        t1 = time.time()
        a = lambda t0, t1: (t1 - t0) / 60
        logger.info(
            'The total running time of function "%s()" is: %.3f minutes' %
            (function.__name__, (a(t0, t1))))
        time.sleep(1)
        return result

    return wrap_function


def clear_fromTML(xlspath):

    file = xlspath
    logger.debug('xlspath is: {}'.format(xlspath))
    rb = xlrd.open_workbook(file, formatting_info=True)
    wb = copy(rb)
    sheet = wb.get_sheet(0)

    for y in range(1, 99):
        try:
            for x in range(3, 66):
                try:
                    sheet.write(y, x, '')
                except Exception as e:
                    logger.error(e)
        except Exception as e:
            logger.error(e)
            break
    os.remove(file)
    wb.save(file)


def generator_y(xlspath, machinetype):

    data = xlrd.open_workbook(xlspath, formatting_info=True)
    table = data.sheets()[0]
    mtlist = []

    for idy in range(1, 99):
        try:
            try:
                cell = table.cell(idy, 2).value
                cell = ''.join(str(int(cell)).split())
                logger.debug('mtcell is: {}'.format(cell))
                mts1 = cell
                if mts1:
                    logger.debug('mts1 is: {}'.format(mts1))
                    mtlist.append(mts1)
                else:
                    logger.error('no mts1!')
                    time.sleep(5)
            except Exception as e:
                if not 'invalid literal for int() with base 10:' in str(e):
                    logger.error(e)
                mts2 = cell
                mts2 = ''.join(cell.split())
                if mts2:
                    logger.debug('mts2 is: {}'.format(mts2))
                    mtlist.append(mts2)
        except Exception as e:
            logger.error(e)
            break
    logger.debug('mtlist is: {}'.format(mtlist))
    dictY = {}
    for value, key in enumerate(mtlist, 1):
        dictY[key] = value
    logger.debug('dictY is: {}'.format(dictY))

    for completekey in dictY.keys():
        try:
            logger.debug('completekey is: {}'.format(completekey))
        except Exception as e:
            logger.error(e)
            pass
        if machinetype.upper() in completekey.upper():
            y = dictY[completekey]
            logger.debug('machinetype is: {}'.format(machinetype))

            return y
    else:
        logger.warning(
            'Base excel no this machinetype: ({}), so invalid machinetype!'.
            format(machinetype))


def generator_x(xlspath, ostype):

    data = xlrd.open_workbook(xlspath, formatting_info=True)
    table = data.sheets()[0]
    oslist = []

    for idx in range(3, 66):
        try:
            cell = table.cell(0, idx).value
            if cell:
                logger.debug('oscell is: {}'.format(cell))
                oslist.append(cell)
        except Exception as e:
            logger.error(e)
            break
    logger.debug('oslist is: {}'.format(oslist))
    dictX = {}

    for value, key in enumerate(oslist, 3):
        dictX[key] = value
    logger.debug('dictX is: {}'.format(dictX))

    for osT in dictX.keys():
        if ostype.upper() == osT.upper():
            x = dictX[osT]
            logger.debug('ostype is: {}'.format(ostype))

            return x
    else:
        logger.warning(
            'Base excel no this ostype: ({}), so invalid ostype!'.format(
                ostype))


class extrawrite():
    def __init__(self, xlspath, eachxfree, Extralist):

        self.xlspath = xlspath
        self.eachxfree = eachxfree
        self.Extralist = Extralist

    def do(self, mtostype='ExtraMT'):

        file = self.xlspath
        rb = xlrd.open_workbook(file, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)

        if len(self.Extralist):
            sheet.write(0, self.eachxfree, mtostype)

            for order, extra in enumerate(self.Extralist, 1):
                logger.debug('{} class is: {}'.format(extra, type(extra)))
                sheet.write(order, self.eachxfree, extra)
                logger.info(
                    '-------- write content is ({0}, {1}, {2}) --------'.
                    format(order, self.eachxfree, extra))
                time.sleep(0.5)
                logger.info(
                    '-------- write the {0}th tml {2}: ({1}) successfully! --------'.
                    format(order, extra, mtostype))
            try:
                os.remove(file)
                wb.save(file)
                logger.debug(
                    '-------- save {} successfully! --------'.format(mtostype))
            except Exception as e:
                logger.error(e)
                logger.info('file is: {}'.format(file))
                logger.error(
                    '-------- save {} failure! --------'.format(mtostype))
                time.sleep(5)
            logger.info(
                '-------- Write {} finished! --------'.format(mtostype))
            logger.info(
                '**************************************************************************************************'
            )
            time.sleep(3)


class extraxfree():
    def __init__(self, xlspath):
        self.xlspath = xlspath

    def caculator(self):

        data = xlrd.open_workbook(self.xlspath, formatting_info=True)
        table = data.sheets()[0]
        oslist = []

        for idx in range(3, 66):
            try:
                cell = table.cell(0, idx).value
                if cell:
                    oslist.append(cell)
            except Exception as e:
                logger.error(e)
                break
        logger.debug('oslist is: {}'.format(oslist))
        xfree = len(oslist) + 4
        logger.debug('xfree is: {}'.format(xfree))

        return xfree


class extrashow():
    def __init__(self, ExtraMT, ExtraOS):

        self.ExtraMT = ExtraMT
        self.ExtraOS = ExtraOS

    def do(self):

        if len(self.ExtraMT) or len(self.ExtraOS):
            logger.info(
                '**************************************** PLEASE NOTICE !!! ***************************************'
            )
        if len(self.ExtraMT):
            logger.info('-------- TML ExtraMT count is: {} --------'.format(
                len(self.ExtraMT)))
            logger.info('-------- TML ExtraMT is: {} --------------'.format(
                self.ExtraMT))
        if len(self.ExtraOS):
            logger.info('-------- TML ExtraOS count is: {} --------'.format(
                len(self.ExtraOS)))
            logger.info('-------- TML ExtraOS is: {} --------------'.format(
                self.ExtraOS))
        if len(self.ExtraMT) or len(self.ExtraOS):
            logger.info(
                '**************************************************************************************************'
            )
            time.sleep(5)


@function_timer
def generator_yx_write_extra(xlspath):

    tlist = []
    ExtraMT = []
    ExtraOS = []

    for order, MTOSmatch in enumerate(match_one_one(), 1):
        logger.info(
            '##########################################################################'
        )
        logger.info('The %-6s %-10s is: %s' % ('{}th'.format(order),
                                               'MTOSmatch', MTOSmatch))
        machinetype = MTOSmatch[0]
        ostype = MTOSmatch[1]
        y = generator_y(xlspath, machinetype)
        logger.debug('y is: {}'.format(y))
        if not y:
            if not machinetype in ExtraMT:
                ExtraMT.append(machinetype)
        x = generator_x(xlspath, ostype)
        logger.debug('x is: {}'.format(x))
        if not x:
            if not ostype in ExtraOS:
                ExtraOS.append(ostype)
        touple = (y, x)
        logger.info('The %-6s %-10s is: %s' % ('{}th'.format(order),
                                               'coordinate', touple))
        tlist.append(touple)
    # if __name__ == '__main__':
    extrashow(ExtraMT, ExtraOS).do()
    xfree = extraxfree(xlspath).caculator()
    extrawrite(xlspath, xfree, ExtraMT).do()
    extrawrite(xlspath, xfree + 1, ExtraOS).do('ExtraOS')

    return tlist


def write_into_fromTML(xlspath):

    tlist = generator_yx_write_extra(xlspath)
    file = xlspath
    rb = xlrd.open_workbook(file, formatting_info=True)
    wb = copy(rb)
    sheet = wb.get_sheet(0)
    fromTmllist = []
    logger.info(
        '-------------- Start to collect and write all match MTOS now! --------------'
    )
    time.sleep(3)
    for order, coordinates in enumerate(tlist, 1):
        logger.debug('coordinates is: {}'.format(coordinates))
        y2 = coordinates[0]
        x2 = coordinates[1]
        t3 = (y2, x2)
        logger.debug('write coordinates t3 is: {}'.format(t3))
        try:
            t4 = (y2 + 1, x2 + 1)
            if t4 not in fromTmllist:
                fromTmllist.append(t4)
        except Exception as e:
            logger.debug(
                "invalid tml element! don't need to add into fromTmllist!")
            logger.info('Skip the {}th invalid coordinates: {}!'.format(
                order, t3))
            if str(
                    e
            ) != "unsupported operand type(s) for +: 'NoneType' and 'int'":
                logger.error(e)
            continue  # no break
        try:
            sheet.write(y2, x2, 'X')
            order = '{}{}'.format(order, 'th')
            logger.info(
                '%-20s write the %-6s match coordinates %-12s %-20s Y(^_^)Y' %
                ('Y(^_^)Y', order, str(coordinates), 'successfully!'))
            logger.info(
                '*************************************************************************************************'
            )
        except Exception as e:
            logger.error(e)
            mark1 = 'XXXXXXX'
            mark2 = 'failure!'
            logger.warning(
                '%-20s write the %-6s match coordinates %-12s %-20s %s' %
                (mark1, '{}{}'.format(order, 'th'), str(coordinates), mark2,
                 mark1))
            time.sleep(5)
            logger.info(
                'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
            )

    os.remove(file)
    wb.save(file)
    logger.info(
        '------------ write all machinetype and ostype finished! ------------')
    logger.info(
        '--------------------------------------------------------------------------------------------------------'
    )
    time.sleep(3)
    fromTmllist.sort()
    logger.info('<<<< fromTmllist element count is: ({}) >>>>'.format(
        len(fromTmllist)))
    time.sleep(1)
    if len(fromTmllist):
        logger.debug('fromTMLlist is: {}'.format(fromTmllist))
        time.sleep(1)

    return fromTmllist


def collect_machinetype_fromTML():

    machinetypelist = []

    for order, f in enumerate(generator_tmlpathlist(), 1):
        try:
            dom = xml.dom.minidom.parse(f)
        except Exception as e:
            logger.error(e)
            logger.warning('Invalid file, maybe not a machinetype tml file!')
            time.sleep(5)
            continue
        root = dom.documentElement
        value = root.getElementsByTagName('MachineTypeModel')

        for n in range(0, 666):
            try:
                if value:
                    value1 = value[n]
                    machinetype = value1.firstChild.data
                    logger.info(machinetype)
                    machinetypelist.append(machinetype)
                else:
                    logger.warning('No machinetype found!')
                    break
            except Exception as e:
                if str(e) != 'list index out of range':
                    logger.error(e)
                break
        logger.info(
            '-------- The {}th tml machinetype collect finished now! --------'.
            format(order))
        time.sleep(1)
    if not len(machinetypelist):
        Error = 'No type collected! please notice!'
        logger.info(Error)
        raise Exception(Error)
    else:
        logger.info('machinetype count is: {}'.format(len(machinetypelist)))
        time.sleep(1)
        logger.debug('machinetypelist is: {}'.format(machinetypelist))
        logger.info('-------- Collect all machinetype finished now! --------')
        time.sleep(1)

        return machinetypelist


class operasys():
    def __init__(self, OpSys, OpSys1, OpSysMany):
        self.OpSys = OpSys
        self.OpSys1 = OpSys1
        self.OpSysMany = OpSysMany

    def collect(self, arch=6):

        self.OpSys = '{}{}'.format(arch, self.OpSys)
        try:
            OpSys3 = str(self.OpSysMany).rsplit(' ')[2]
            if self.OpSys1 == 'W':
                if OpSys3.upper() == 'R2':
                    self.OpSys = '{}{}'.format(self.OpSys, OpSys3)
                    logger.info(self.OpSys)
                elif OpSys3.lower() == 'version':
                    self.OpSys = 'W{}'.format(
                        str(self.OpSysMany).rsplit(' ')[3])
                    logger.info(self.OpSys)
                else:
                    logger.info(self.OpSys)
            elif self.OpSys1.lower() == 'vmware':
                self.OpSys = '{}VMESXI{}'.format(arch, OpSys3)
                logger.info(self.OpSys)
            else:
                logger.info(self.OpSys)
        except Exception as e:
            logger.info(self.OpSys)
            if str(e) != 'list index out of range':
                logger.error(e)

        return self.OpSys


@function_timer
def collect_ostype_fromTML():

    ostypelist = []

    for order, f in enumerate(generator_tmlpathlist(), 1):
        try:
            dom = xml.dom.minidom.parse(f)
        except Exception as e:
            logger.error(e)
            logger.warning('Invalid file, maybe not a ostype tml file!')
            time.sleep(5)
            continue
        root = dom.documentElement
        value = root.getElementsByTagName('OpSys')
        valuex = root.getElementsByTagName('ProcessorType')
        for n in range(0, 666):
            try:
                if value:
                    value1 = value[n]
                    OpSysMany = value1.firstChild.data
                    logger.info(OpSysMany)
                    OpSys1 = str(OpSysMany).rsplit(' ')[0]
                    if OpSys1.upper() == 'WINDOWS':
                        OpSys1 = 'W'
                    OpSys2 = str(OpSysMany).rsplit(' ')[1]
                    OpSys = '{}{}'.format(OpSys1, OpSys2)
                else:
                    logger.warning('No OpSys found!')
                    break
                if valuex:
                    value2 = valuex[n]
                    ProcessorType = value2.firstChild.data
                    logger.info(ProcessorType)
                    if ProcessorType.lower() == 'x64':
                        ostypelist.append(
                            operasys(OpSys, OpSys1, OpSysMany).collect())
                    elif ProcessorType.lower(
                    ) == 'x86' or ProcessorType == 'x32':
                        ostypelist.append(
                            operasys(OpSys, OpSys1, OpSysMany).collect(3))
                    else:
                        logger.error(
                            'Wrong ProcessorType: ({})!'.format(ProcessorType))
                        errorType = "Wrong ProcessorType: arch={}".format(
                            ProcessorType)
                        ostypelist.append(errorType)
                else:
                    logger.warning('No ProcessorType found!')
                    break
            except Exception as e:
                if str(e) != 'list index out of range':
                    logger.error(e)
                break
        logger.info(
            '-------- The {}th tml ostype collect finished now! --------'.
            format(order))
        time.sleep(1)

    if not len(ostypelist):
        Error = 'No type collected! please notice!'
        logger.info(Error)
        raise Exception(Error)
    else:
        logger.info('ostypelist count is: {}'.format(len(ostypelist)))
        time.sleep(1)
        logger.debug('ostypelist is: {}'.format(ostypelist))
        logger.info('-------- Collect all ostype finished now! --------')
        time.sleep(1)

    return ostypelist


def match_one_one():

    machinetypelist = collect_machinetype_fromTML()
    ostypelist = collect_ostype_fromTML()
    if len(machinetypelist) == len(ostypelist):
        mtos = zip(machinetypelist, ostypelist)

        return mtos
    else:
        Error = 'oslist count not equal machinelist count!'
        logger.error(Error)

        raise Exception(Error)


def web_link():

    weblink = {
        'loginPage':
        'http://onestop.labs.lenovo.com',
        'urlRedLogin':
        'http://rsl-ossweb20.labs.lenovo.com:9084/OssWeb/RedLogon.do',
        'urlSearch':
        'http://rsl-ossweb20.labs.lenovo.com:9084/OssWeb/newSearch.do'
    }

    for links in weblink.values():

        yield links


def get_session_post(s):

    attempts = -1
    success = False
    while attempts < 150 and not success:
        postdata = {
            "userId": "yanjy2@lenovo.com",
            "password": "1203",
            "submit": "Log In"
        }
        logger.debug('postdata is: %s' % postdata)
        for number, link in enumerate(web_link(), 1):
            logger.debug('Post the {0}th {1}'.format(number, link))
            try:
                s.post(link, postdata)
                success = True
            except Exception as e:
                logger.error(e)
                logger.error(
                    'Post failure, maybe network off-line or OSS crash!!!')
                time.sleep(5)
                attempts += 1
                if attempts < 150:
                    logger.info('****** Now retry the %dth time ******' %
                                (attempts + 1))
                break
    if not success:
        logger.info('#### Please try again later! ####')
        sys.exit()


def make_tmlfolder(link, formid):

    if 'lnvgy_utl_lxce' in link.lower():
        if 'ux' in link.lower() and 'anyos' in link.lower():
            formidNew = '{}_{}'.format(formid, 'UX_For_BoMC')
        elif '_bomc' in link.lower():
            formidNew = '{}_{}'.format(formid, 'BoMC')
        elif '_onecli' in link.lower():
            formidNew = '{}_{}'.format(formid, 'OneCLI')
        elif '_ux' in link.lower() and not 'anyos' in link.lower():
            formidNew = '{}_{}'.format(formid, 'OneGUI')
        else:
            formidNew = '{}_{}'.format(formid, 'otherForm')
    else:
        if 'boot' in link.lower() and 'bomc' in link.lower():
            formidNew = '{}_{}'.format(formid, 'SaLIE')
        elif 'boot' in link.lower() and 'tools' in link.lower(
        ) and '7.4' in link:
            formidNew = '{}_{}'.format(formid, 'MCP')
        elif '_dsa_' in link.lower():
            formidNew = '{}_{}'.format(formid, 'DSA')
        elif '_asu_' in link.lower():
            if not 'rpm' in link.lower():
                formidNew = '{}_{}'.format(formid, 'ASU')
            else:
                formidNew = '{}_{}'.format(formid, 'RPM_ASU')
        elif '_uxspi_' in link.lower():
            formidNew = '{}_{}'.format(formid, 'UXSPI')
        else:
            formidNew = '{}_{}'.format(formid, 'otherForm')

    return formidNew


@function_timer
def download(s, formid):

    formidNew = 'none'
    for num, link in enumerate(get_urlLinks(s, formid), 1):
        logger.debug('link is: %s' % link)
        linkYes = '{}{}'.format(
            'http://rsl-ossweb20.labs.lenovo.com:9084/OssWeb/', link)
        logger.debug('linkyes is: %s' % linkYes)
        if not '=' in link:
            logger.warning('The invalid link is: %s!' % link)
            continue
        filename = str(link).split('=')[1]
        logger.debug('filename is: %s' % filename)
        if '//' in filename and 'http' in filename.lower():
            logger.warning('#### Invalid file name: %s ####.', filename)
            continue
        formidNew = make_tmlfolder(link, formid)
        path = os.path.join(sys.path[0], formidNew)
        logger.debug('path is: %s' % path)
        pathname = os.path.join(path, filename)
        logger.debug('pathname is: %s' % pathname)
        if not os.path.exists(path):
            os.makedirs(path)
        logger.info("Downloading the {0}th file: {1}".format(num, pathname))
        response = s.get(linkYes, stream=True)
        chunk_size = 1024 * 50
        size = response.headers.get('content-length')
        logger.debug('size is: %s' % size)
        if size:
            content_size = int(size)
            logger.debug('if size is: {}'.format(int(size)))
        else:
            logger.warning('#### Invalid link! ####')
            continue
        widgets = [
            'Downloaded: ',
            Percentage(), ' ',
            Bar('*'), ' ',
            Timer(), '  ',
            ETA(), ' ',
            FileTransferSpeed()
        ]
        with open(pathname, mode="wb") as f:
            with ProgressBar(widgets=widgets, max_value=content_size) as pbar:
                datalen = 0
                for data in response.iter_content(chunk_size=chunk_size):
                    datalen += len(data)
                    logger.debug('datalen is: {}'.format(datalen))
                    f.write(data)
                    pbar.update(datalen)
        if content_size < 28888888:
            continue
        else:
            logger.debug(
                "File's size is %d, too big to download much time. Maybe next file will be refused by OSS so post again!"
                % content_size)
            get_session_post(s)

    if formidNew != 'none':
        return formidNew
    else:
        return 'none'


def collect_formid():

    n = 0
    formids = []
    while True:
        formid = input('Please input formid now: ')
        if formid == '' and len(formids):
            break
        else:
            if formid.isdigit() and not formid in formids and int(formid) > 0:
                formids.append(formid)
                n += 1
                logger.info('Add the {}th formid successfully!'.format(n))
                logger.info(
                    '------ If no formid need input, you can press "Enter" to download! ------'
                )
            else:
                logger.error(
                    "Input error, formid must be integer (great than 0!), different and formids isn't empty!!!"
                )

    logger.info('Formid list is: {}'.format(formids))
    logger.info('-------- Formid count is: {} --------'.format(len(formids)))
    time.sleep(3)

    return formids


def get_urlLinks(s, formid):

    urlFormID = '{}{}'.format(
        'http://rsl-ossweb20.labs.lenovo.com:9084/OssWeb/DisplayOssForm.do?formId=',
        formid)
    logger.debug('urlFormID is: %s' % urlFormID)
    html = s.get(urlFormID)
    logger.debug('html is: %s' % html)
    soup = BeautifulSoup(html.text, 'lxml')
    soup1 = soup.decode('utf-8')
    logger.debug('soup is: %s' % soup1)
    divs = soup.find_all('tr', class_='oss-tbody')
    for div in divs:
        div1 = div.encode('utf-8')
        logger.debug('div1 is: %s' % div1)
        if div.a:
            logger.debug('div.a is: %s' % div.a)
            links = div.a.get('href')
            if '.tml' in links.lower():

                yield links


def iter_and_download_formid(s, formids):

    delete_otherForm()

    for order, formid in enumerate(formids, 1):
        get_session_post(s)
        logger.info(
            '###############################################################################'
        )
        logger.info(
            '###############################################################################'
        )
        logger.info(
            '##                                                                           ##'
        )
        logger.info('%-15s Begin to download the %-7s formid: %-22s %s' %
                    ('##', '{}{}'.format(order, 'th'), [formid], '##'))
        logger.info(
            '##                                                                           ##'
        )
        logger.info(
            '###############################################################################'
        )
        logger.info(
            '###############################################################################'
        )
        time.sleep(3)

        formidNew = download(s, formid)

        logger.info('formidNew is: {}'.format([formidNew]))
        try:
            delete_old_form(formidNew)
            show_download_result(formid)
        except Exception as e:
            if str(e) == 'invalid formid, so no formidNew generator':
                show_download_result(formid)
                continue
            else:
                logger.error(e)
                break


def delete_otherForm():

    path = sys.path[0]
    foldernameList = os.listdir(path)
    logger.debug('foldernameList is: {}'.format(foldernameList))

    if foldernameList:
        for tmlfolder in foldernameList:
            if not '_py' in tmlfolder.lower(
            ) and not 'logs' in tmlfolder.lower() and str(tmlfolder).rsplit(
                    '_')[0].isdigit() and str(tmlfolder).rsplit(
                        '_')[1] and 'otherform' in tmlfolder.lower():
                if tmlfolder:
                    logger.info(
                        'Now delete old otherform ({})!'.format(tmlfolder))
                    dirpath = os.path.join(path, tmlfolder)
                    shutil.rmtree(dirpath, True)
                    logger.info(
                        'Delete old folder ({}) finished!'.format(tmlfolder))


def delete_old_form(formidNew):

    if formidNew == 'none':
        Error = 'invalid formid, so no formidNew generator'
        raise Exception(Error)

    path = sys.path[0]
    foldernameList = os.listdir(path)
    logger.debug('foldernameList is: {}'.format(foldernameList))

    if foldernameList:
        for tmlfolder in foldernameList:
            if not '_py' in tmlfolder.lower(
            ) and not 'logs' in tmlfolder.lower() and str(tmlfolder).rsplit(
                    '_')[0].isdigit() and str(tmlfolder).rsplit(
                        '_')[1] and not 'otherform' in tmlfolder.lower():
                toolname = str(tmlfolder).rsplit('_')[1]
                logger.debug(tmlfolder)
                logger.debug(toolname)
                formidtools = str(formidNew).rsplit('_')[1].lower()
                if formidtools == toolname.lower() and formidNew != tmlfolder:
                    if formidtools != 'onegui' and formidtools != 'ux' and formidtools != 'rpm' and formidtools != 'asu':
                        logger.info(
                            'Now delete old tmlfolder ({})!'.format(tmlfolder))
                        dirpath = os.path.join(path, tmlfolder)
                        shutil.rmtree(dirpath, True)
                        logger.info('Delete old folder ({}) finished!'.format(
                            tmlfolder))
                    else:
                        if formidtools == 'onegui':
                            del1to2(tmlfolder, path, 'UX_For_BoMC').do()
                        elif formidtools == 'ux':
                            del1to2(tmlfolder, path, 'OneGUI').do()
                        elif formidtools == 'rpm':
                            del1to2(tmlfolder, path, 'ASU').do()
                        else:
                            del1to2(tmlfolder, path, 'RPM_ASU').do()


class del1to2():
    def __init__(self, tmlfolder, path, tml_toolname):

        self.tmlfolder = tmlfolder
        self.path = path
        self.tml_toolname = tml_toolname

    def do(self):

        logger.info('Now delete old tmlfolder ({})!'.format(self.tmlfolder))
        dirpath = os.path.join(self.path, self.tmlfolder)
        shutil.rmtree(dirpath, True)
        logger.info('Delete old folder ({}) finished!'.format(self.tmlfolder))
        tml_foldername = '{}_{}'.format(
            str(self.tmlfolder).rsplit('_')[0], self.tml_toolname)
        dir2path = os.path.join(self.path, tml_foldername)
        if os.path.exists(dir2path):
            logger.info(
                'Now delete old tmlfolder ({})!'.format(tml_foldername))
            shutil.rmtree(dir2path, True)
            logger.info(
                'Delete old folder ({}) finished!'.format(tml_foldername))


def show_download_result(formid):

    path = sys.path[0]
    foldernameList = os.listdir(path)
    folderlist = []

    for tmlfolder in foldernameList:
        if not '_py' in tmlfolder.lower() and not 'logs' in tmlfolder.lower(
        ) and str(tmlfolder).rsplit('_')[0].isdigit():
            folderlist.append(str(tmlfolder).rsplit('_')[0])
    logger.debug('folderlist is: {}'.format(folderlist))

    if not formid in folderlist:
        logger.warning('[{}] download failure！'.format(formid))
        logger.warning('[{}] is invalid formid!!!'.format(formid))
        logger.warning('please confirm the formid is rigth!')
        time.sleep(3)
    else:
        logger.info('[{}] download finished！'.format(formid))


class compare():
    def __init__(self, xlspath, comparelist, comparetype):

        self.xlspath = xlspath
        self.comparelist = comparelist
        self.comparetype = comparetype

    def write(self):

        file = self.xlspath
        rb = xlrd.open_workbook(file, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)

        for order, coordinates in enumerate(self.comparelist, 1):
            y8 = coordinates[0]
            x8 = coordinates[1]
            t8 = (y8, x8)
            logger.debug('{} coordinates t8 is: {}'.format(
                self.comparetype, t8))
            sheet.write(y8 - 1, x8 - 1, '')
            sheet.write(y8 - 1, x8 - 1, self.comparetype)
            logger.debug(
                ' ---- write {0}th coordinates {1} successfully! ---- '.format(
                    order, coordinates))
        os.remove(file)
        wb.save(file)

        if len(self.comparelist):
            logger.info('-------- waiting for writing {}! --------'.format(
                self.comparetype))
            time.sleep(2)
            logger.info(
                '-------- The {} elements write "{}" in the compare_result excel! --------'.
                format(self.comparetype.upper(), self.comparetype))
            time.sleep(1)


class compareshow():
    def __init__(self, Extralist, Missinglist):

        self.Extralist = Extralist
        self.Missinglist = Missinglist

    def do(self):

        logger.info(
            '*****************************************************************************************************'
        )
        if len(self.Extralist):
            logger.debug('Extralist is: {}'.format(self.Extralist))
        a = len(self.Extralist)
        if a:
            logger.info(
                '********* There are {} extra elements in the tml file that are not needed! need to delete! ********'.
                format(a))
        if len(self.Missinglist):
            logger.debug('Missinglist is: {}'.format(self.Missinglist))
        b = len(self.Missinglist)
        if b:
            logger.info(
                '********* There are {} missing elements in the tml file that are needed! need to add! ***********'.
                format(b))
        logger.info(
            '-------------- Please view coordinates in the excel file! -------------------------------------------'
        )
        logger.info(
            '-------------- Please reference base excel file and fromTML excel file! -----------------------------'
        )
        if not a and not b:
            logger.info(
                '-------------- All exitsed excel info are same! But you must check ExtraOS and ExtraMT! -------------'
            )
        logger.info(
            '*****************************************************************************************************'
        )
        time.sleep(3)


def compare_result(xlspath, **kwargs):

    Missinglist = []
    Extralist = []

    logger.debug(kwargs)
    if 'fromTmllist' in kwargs and 'baselist' in kwargs:

        for i in kwargs.get('baselist'):
            if i not in kwargs.get('fromTmllist'):
                Missinglist.append(i)

        for j in kwargs.get('fromTmllist'):
            if j not in kwargs.get('baselist'):
                Extralist.append(j)

        # if __name__ == '__main__':

        compare(xlspath, Missinglist, 'Missing').write()
        compare(xlspath, Extralist, 'Extra').write()
        compareshow(Extralist, Missinglist).do()


def show_logfolderpath():

    time.sleep(1)
    logger.info('Saving logfile now, please wait.............')
    time.sleep(1)
    logger.info(
        '***************************************************************************'
    )
    logger.info('<<<<  Logfile is saved in: {0} >>>>'.format(G_logfolderpath))
    logger.info(
        '***************************************************************************'
    )
    time.sleep(1)


def create_logfile_path():

    G_logfolderpath = os.path.join(sys.path[0], 'Logs')
    if not os.path.exists(G_logfolderpath):
        os.makedirs(G_logfolderpath)
    tname = time.strftime("%Y%m%d%H%M%S", time.localtime())
    filename = '{}{}{}'.format('base_and_tml_compare_result_', tname, '.LOG')
    logfilepath = os.path.join(G_logfolderpath, filename)

    return G_logfolderpath, logfilepath


def create_logger_func(logfilepath):

    logger = logging.getLogger("name")
    logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler(logfilepath, mode='w')
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    file_formatter = logging.Formatter(
        "%(asctime)s - [%(funcName)s()] - [line:%(lineno)d] - %(levelname)s: %(message)s"
    )
    console_formatter = logging.Formatter('%(message)s')
    file_handler.setFormatter(file_formatter)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    file_handler.setLevel(logging.INFO)

    try:
        options, args = getopt.getopt(sys.argv[1:], 'HVF:D:L:A:', [
            'authored=', "help", 'version', "formid=", 'download=',
            'log-level='
        ])
        if len(options) or len(args):
            logger.debug('options is: {0}, args is: {1}'.format(options, args))
    except getopt.GetoptError as e:
        logger.error(e)
        quit()

    for arg in args:
        if arg:
            logger.info('invalid args: ({})!'.format(arg))
            quit()

    for name, value in options:
        if name in ('-L', '--log-level'):
            if value == 'debug':
                file_handler.setLevel(logging.DEBUG)
                logger.info(
                    'You specificate option ({}={}), now it will collect <{}> level logfile!'.
                    format(name, value, value))
                time.sleep(1)
            elif value == 'warn':
                file_handler.setLevel(logging.WARN)
                logger.info(
                    'You specificate option ({}={}), now it will collect <{}> level logfile!'.
                    format(name, value, value))
            else:
                logger.info('args: ({}) value: ({}), value error!'.format(
                    name, value))
                quit()
        if name in ("-H", "--help"):
            if len(options) != 1:
                logger.info(
                    "Don't use these args: ({}) at the same time!".format(
                        options))
                quit()
            else:
                logger.info(
                    '-H; --help\n-V; --version\n-F; --formid(No.)\n-D; --download(YES/NO)\n-L; --log-level(debug/warn)\n-A; --authored(by/time)'
                )
                quit()
        if name in ('-V', '--version'):
            if len(options) != 1:
                logger.info(
                    "Don't use these args: ({}) at the same time!".format(
                        options))
                quit()
            else:
                logger.info('Version: 1.0.0')
                quit()
        if name in ('-A', '--authored'):
            if len(options) != 1:
                logger.info(
                    "Don't use these args: ({}) at the same time!".format(
                        options))
                quit()
            else:
                if value == 'time':
                    logger.info('It was authored in May 2018!')
                    quit()
                elif value == 'by':
                    logger.info('It was authored by Isaac.Yan!')
                    quit()
                else:
                    logger.info('args: ({}) value: ({}), value error!'.format(
                        name, value))
                    quit()
        if name in ('-F', '--formid'):
            logger.info('formid is: {}'.format(value))
            quit()
        if name in ('-D', '--download'):
            if value == 'yes' or value == 'no':
                if value == 'yes':
                    logger.info('way = "1"')
                    quit()
                else:
                    logger.info('way = "2"')
                    quit()
            else:
                logger.info('wrong value: {}'.format(value))
                quit()
    if len(options) or len(args):
        logger.info('Analysis opts and args finished!')
        time.sleep(1)

    return logger


def run_choice_func():

    while True:
        way = input(
            '-------- Could you need to download form? --------\nplease input (YES or NO ?): '
        )
        logger.debug('Input way is: {}'.format(way))
        logger.debug('{ YES(need to download form), NO(no need) }')
        if not way:
            logger.info('No choice!')
            continue
        if way.upper() == 'YES' or way.upper() == 'NO':
            break
        else:
            logger.info('Wrong input!')
    if way.upper() == 'YES':
        logger.info('-------- Download form at the first! --------')
        time.sleep(1)
        run_haveget_func()
    else:
        logger.info(
            '-------- Compare base excel and OSS tml directly now! --------')
        time.sleep(1)
        run_noget_func()


def copy_baseExcel_create_bak_iter_base_runCompare():

    xlsList = []
    path = sys.path[0]
    logger.debug('script path is: {}'.format(path))
    filenameList = os.listdir(path)
    logger.debug('Excel filenameList is: {}'.format(filenameList))

    for name in filenameList:
        if not 'fromTML'.lower() in name.lower() and not 'bak'.lower(
        ) in name.lower() and '.xls'.lower() in name.lower() and '_' in name:
            xlsList.append(name)
    if not len(xlsList):
        logger.critical(
            'No base excel file found! please move into this folder ({})'.
            format(path))
        show_logfolderpath()
        quit()
    else:
        global g_excel_file_name
        resultList = []

        for excelorder, g_excel_file_name in enumerate(xlsList, 1):
            logger.debug('base name is: {}'.format(g_excel_file_name))
            noEnd = str(g_excel_file_name).rsplit('.')[0]
            logger.debug('excel full name is: {}'.format(
                str(g_excel_file_name).rsplit('.')))
            logger.debug('noEnd is: {}'.format(noEnd))
            file1 = os.path.join(path, g_excel_file_name)
            logger.debug('file1 is: {}'.format(file1))
            file2 = '{}-fromTML-compare-result.xls'.format(
                os.path.join(path, noEnd))
            logger.debug('file2 is: {}'.format(file2))
            file3 = '{}-bak.xls'.format(os.path.join(path, noEnd))
            logger.debug('file3 is: {}'.format(file3))
            shutil.copyfile(file1, file2)
            shutil.copyfile(file1, file3)

            runcompare(file2, resultList).show_excel_compare_order(excelorder)

            xlspath = file2
            clear_fromTML(xlspath)
            try:
                compare_result(
                    xlspath,
                    fromTmllist=write_into_fromTML(xlspath),
                    baselist=read_baseXLS(file1))
                resultList.append(
                    'Compare %-35s => Success' % g_excel_file_name)
            except Exception as e:
                resultList.append(
                    'Compare %-35s => Failure' % g_excel_file_name)
                if str(
                        e
                ) == 'no corresponding tml folder found! please notice!':
                    logger.error(
                        'No corresponding tml folder!! please confirm whether it is right formid!'
                    )
                    time.sleep(5)
                    continue
                elif str(
                        e
                ) == 'no tml found in this tml folder! please notice!':
                    time.sleep(5)
                    continue
                elif str(e) == 'No type collected! please notice!':
                    time.sleep(5)
                    continue
                elif str(e) == 'oslist count not equal machinelist count!':
                    time.sleep(5)
                    continue
                else:
                    logger.critical(
                        'One of three func appear error: {}'.format(e))
                    time.sleep(5)
                    show_logfolderpath()
                    quit()

        runcompare(file2, resultList).show_each_compare_result()


class runcompare():
    def __init__(self, file2, resultList):

        self.file2 = file2
        self.resultList = resultList

    def show_excel_compare_order(self, excelorder):

        logger.info(
            '#####################################################################################################'
        )
        logger.info(
            '#####################################################################################################'
        )
        logger.info(
            '##                                                                                                 ##'
        )
        logger.info(
            '%-16s Start to compare the %-4s baseExcel: %-44s %s' %
            ('##', '{}{}'.format(excelorder, 'th'), [g_excel_file_name], '##'))
        logger.info(
            '##                                                                                                 ##'
        )
        logger.info(
            '#####################################################################################################'
        )
        logger.info(
            '#####################################################################################################'
        )
        time.sleep(2)
        logger.info('The {}th compare result is saved in: {}'.format(
            excelorder, self.file2))
        time.sleep(2)

    def show_each_compare_result(self):

        logger.debug('resultList is: {}'.format(self.resultList))

        for result in self.resultList:
            logger.info(result)
            time.sleep(1)

        if not 'Fail' in str(self.resultList):
            logger.info(
                'All of Tools MTOS match relationship compare finished and successfully now!'
            )
            time.sleep(3)
        elif 'Fail' in str(self.resultList) and 'Success' in str(
                self.resultList):
            logger.warning(
                'Not all of Tools MTOS match relationship compare successfully!'
            )
            time.sleep(3)
        else:
            logger.warning(
                'All of Tools MTOS match relationship compare failure!')
            time.sleep(3)


def read_baseXLS(file1):

    logger.info('read base excel is: {}'.format(file1))
    time.sleep(3)
    path = sys.path[0]
    baseXlspath = os.path.join(path, file1)
    logger.debug('baseXlspath is: %s' % baseXlspath)
    data = xlrd.open_workbook(baseXlspath, formatting_info=True)
    table = data.sheets()[0]
    baselist = []

    for y in range(1, 99):
        try:
            for x in range(3, 66):
                try:
                    cell = table.cell(y, x).value
                    try:
                        if cell.upper() == 'X':
                            t5 = (y + 1, x + 1)
                            baselist.append(t5)
                    except Exception as e:
                        logger.error(e)
                        logger.debug('add fail!')
                        break
                except Exception as e:
                    if str(e) != 'array index out of range' and str(
                            e) != 'list index out of range':
                        logger.error(e)
                    break
        except Exception as e:
            logger.debug(e)
            logger.debug('read base y finished')
            break
    baselist.sort()
    logger.info("<<<< baselist element count is: ({}) >>>>".format(
        len(baselist)))
    time.sleep(1)
    if len(baselist):
        logger.debug('baselist is: {}'.format(baselist))
        time.sleep(1)

    return baselist


def collect_tmlfolder():

    path = sys.path[0]
    foldernameList = os.listdir(path)
    folderlist = []

    for tmlfolder in foldernameList:
        if not '_py' in tmlfolder.lower() and not 'logs' in tmlfolder.lower(
        ) and str(tmlfolder).rsplit('_')[0].isdigit() and str(
                tmlfolder).rsplit('_')[1]:
            folderlist.append(tmlfolder)

    if not len(folderlist):
        logger.critical(
            'At this path: [{}], no tml file folder found!'.format(path))
        logger.critical(
            'Please download form at first and confirm your input formid is valid!'
        )
        time.sleep(3)
        show_logfolderpath()
        quit()
    else:
        logger.debug('tml folderlist is: {}'.format(folderlist))
        logger.debug('tmlFolderList count is: {}'.format(len(folderlist)))
    foldername = match_tml_base(folderlist)

    return foldername


def match_tml_base(folderlist):

    for foldername in folderlist:
        if 'bomc' in g_excel_file_name.lower(
        ) and not 'onegui' in g_excel_file_name.lower(
        ) and 'bomc' in foldername.lower(
        ) and not '_ux_' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'onecli' in g_excel_file_name.lower(
        ) and 'onecli' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'onegui' in g_excel_file_name.lower(
        ) and not 'bomc' in g_excel_file_name.lower(
        ) and 'onegui' in foldername.lower(
        ) and not 'bomc' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'onegui' in g_excel_file_name.lower(
        ) and 'bomc' in g_excel_file_name.lower(
        ) and '_ux_' in foldername.lower() and 'bomc' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'mcp' in g_excel_file_name.lower() and 'mcp' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'salie' in g_excel_file_name.lower(
        ) and 'salie' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'dsa' in g_excel_file_name.lower() and 'dsa' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'asu' in g_excel_file_name.lower(
        ) and not 'rpm' in g_excel_file_name.lower(
        ) and 'asu' in foldername.lower() and not 'rpm' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'rpm' in g_excel_file_name.lower(
        ) and 'asu' in g_excel_file_name.lower() and 'rpm' in foldername.lower(
        ) and 'asu' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername

        if 'uxspi' in g_excel_file_name.lower(
        ) and 'uxspi' in foldername.lower():
            tmlfoldername = foldername
            logger.debug('foldername is: {}'.format(tmlfoldername))

            return foldername
    else:
        Error = 'no corresponding tml folder found! please notice!'

        raise Exception(Error)


def generator_tmlpathlist():

    tmlfoldername = collect_tmlfolder()

    logger.info('tmlfoldername is: {}'.format(tmlfoldername))
    logger.debug('g_excel_file_name is: {}'.format(g_excel_file_name))
    time.sleep(3)

    path = sys.path[0]
    tmlindex = os.path.join(path, tmlfoldername)
    logger.debug('tmlindex is: %s' % tmlindex)
    tmlnamelist = os.listdir(tmlindex)

    if not tmlnamelist:
        logger.warning(
            'No tml file found in this tml file path: {0}'.format(tmlindex))
        Error = 'no tml found in this tml folder! please notice!'
        raise Exception(Error)
    else:
        logger.debug('tmlnamelist is: %s' % tmlnamelist)

        for order, name in enumerate(tmlnamelist, 1):
            tmlpath = os.path.join(tmlindex, name)
            logger.debug('tmlpath is: %s' % tmlpath)
            logger.info('Start to collect the {}th tml file: {}'.format(
                order, tmlpath))
            time.sleep(2)

            yield tmlpath


def run_create_logfile():

    global G_logfolderpath
    global logger
    G_logfolderpath = create_logfile_path()[0]
    logfilepath = create_logfile_path()[1]
    logger = create_logger_func(logfilepath)


def run_get_func():

    s = requests.Session()
    formids = collect_formid()
    iter_and_download_formid(s, formids)


def run_haveget_func():

    run_get_func()
    copy_baseExcel_create_bak_iter_base_runCompare()


def run_noget_func():

    copy_baseExcel_create_bak_iter_base_runCompare()


def run_have_choice_func():

    run_create_logfile()
    run_choice_func()
    show_logfolderpath()


def run_no_choice_func():

    run_create_logfile()
    run_haveget_func()
    show_logfolderpath()


class capture():
    @function_timer
    def run(self):
        try:
            run_have_choice_func()
        except KeyboardInterrupt as e:
            logger.error(e)
            logger.error('"Ctrl + C" keyin, Interrupt!')
            show_logfolderpath()


if __name__ == '__main__':
    capture().run()
