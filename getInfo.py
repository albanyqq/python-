# _*_ coding:utf-8 _*_

import xlwt
import os
import exifread


from PIL import Image
from PIL.ExifTags import TAGS


def get_pic():
    #g = os.walk("G:\\2016卷年鉴来稿照片（原片）\\职业与成人教育\\国家重点中等职业学校\\17北京市实美职业学校20160315原稿")
    g = os.walk('D:\\tupian')
    pics = list()
    for path, d, filelist in g:
        for filename in filelist:
            if filename.endswith('jpg') or filename.endswith('JPG')or filename.endswith('png') or filename.endswith('PNG')or filename.endswith('jpegg') or filename.endswith('JPEG'):
                pics.append(os.path.join(path, filename))
    return pics


def get_exif_data(fname):
    ret = {}
    ret2 = {}
    try:
        img = Image.open(fname)
        if hasattr(img, '_getexif'):
            exifinfo = img._getexif()
            if exifinfo != None:
                for tag, value in exifinfo.items():
                    decoded = TAGS.get(tag, tag)
                    ret[decoded] = value
    except Exception as e:
        print(e)
    if ret:
        try:
            FIELD = 'EXIF DateTimeOriginal'
            f = open(fname, 'rb')
            tags = exifread.process_file(f)
            f.close()
            tags = tags['Image Artist']
            ret2['Artist'] = str(tags)
            ret2['DateTimeOriginal'] = ret['DateTimeOriginal'].split(' ')[0].replace(':', '/')
            if FIELD in tags:
                  new_name = str(tags[FIELD]).replace(':', '').replace(' ', '_')
                  print('=== ', new_name)
            else:
                print('No {} found'.format(FIELD))
        except Exception as e:
            print(e)
            pass
    return ret2


if __name__ == '__main__':
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet1")
    sheet.write(0, 0, '图片名')
    sheet.write(0, 1, '拍摄作者')
    sheet.write(0, 2, '拍摄时间')
    sheet.write(0, 3, '图片路径')
    pics = get_pic()
    s = input("请输入编号:")
    m = input("请输入起始号码:")
    i = int(m)
    for each in pics:
        print(each)
        FIELD = 'EXIF DateTimeOriginal'
        f = open(each, 'rb')
        tags = exifread.process_file(f)
        f.close()
        try:
            time = str(tags[FIELD]).split(' ')[0].replace(':', '/')
        except Exception as e:
            print(e)
            pass
        print( time)
        name = s+'%04d--'% i + each.split('\\')[-1]
        exifs = get_exif_data(each)
        if exifs:
            sheet.write(i, 0, name)
            sheet.write(i, 1, exifs['Artist'])
            sheet.write(i, 2, exifs['DateTimeOriginal'])
            sheet.write(i, 3, each)
        else:
            try:
                sheet.write(i, 0, name)
                sheet.write(i, 2, time)
                sheet.write(i, 3, each)
            except Exception as e:
                print(e)
            pass
        tmp = each.split('\\')
        tmp.pop()
        new_name = '\\'.join(tmp)+'\\'+name
        os.rename(each, new_name)
        i += 1
    workbook.save("shuju.xls")
