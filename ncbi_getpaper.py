# -*- coding: utf-8 -*-

import io
import sys
import urllib
import urllib2
from bs4 import BeautifulSoup
import xlwt
import time

reload(sys)
sys.setdefaultencoding('utf-8')

searchKwd = sys.argv[1]
if len(sys.argv) == 3:
    pageLimited = int(sys.argv[2])/10
else:
    pageLimited = 20

sortMethod = '&sort=[relevance]'

print '\r\nWorking hard to search %s ... ...' % searchKwd
searchUrl = 'https://www.ncbi.nlm.nih.gov/m/pubmed/?term='
paperCount = 0
res = urllib2.urlopen(searchUrl + searchKwd + sortMethod)
soup = BeautifulSoup(res,"html.parser")
book_div = soup.find(attrs={"class":"r"})
if book_div is None:
    print '\r\nSorry, No paper of keyword %s' % searchKwd
    sys.exit(0)
book_a = book_div.findAll(attrs={"rel":"chapter"})
book_span = book_div.findAll(attrs={"class":"aux light_narrow_text"})

pages = soup.find(attrs={"class":"p sml mid"})
pagecount = int(pages.string.strip().split()[-1])

resultCount = soup.find('span', class_='light_narrow_text').get_text()
print('Total return %s of %s' % (resultCount, searchKwd))

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('paper', cell_overwrite_ok=True)
row0 = [u'Paper Titie', u'Author', u'Jour', u'Year']  

sheet.col(0).width = 1224*20
sheet.col(1).width = 300*20
sheet.col(2).width = 350*20
sheet.col(3).width = 100*20

tall_style = xlwt.easyxf('font: name Times New Roman, height 300, bold on', num_format_str='#,##0.00')

for i in range(len(row0)):
    sheet.write(0, i, row0[i], tall_style)

for k in range(len(book_a)):
    link = 'https://www.ncbi.nlm.nih.gov/m/pubmed' + book_a[k]['href'].replace('.', '')
    paperInfo = book_span[k].string.replace('\n', '').split('.')
    sheet.write(paperCount + 1, 1, paperInfo[0]) #author
    sheet.write(paperCount + 1, 2, paperInfo[1].strip()) #Jour
    sheet.write(paperCount + 1, 3, paperInfo[2]) #time
    book_a[k].string = book_a[k].string.replace('"', '')
    mylink = 'HYPERLINK("%s";"%s")' % (link, book_a[k].string)
    sheet.write(paperCount + 1, 0, xlwt.Formula(mylink))
    paperCount += 1
res.close()

if pagecount > pageLimited:
    pagecount = pageLimited


headers = {'User-agent': 'Mozilla/5.0 (Windows NT 6.2; WOW64; rv:22.0) Gecko/20100101 Firefox/22.0'}
if pagecount > 1 :
    for i in range(2, pagecount + 1):
        try:
            url = searchUrl + searchKwd + '&page=' + str(i) + sortMethod
            request = urllib2.Request(url, headers=headers)
            res = urllib2.urlopen(request)
            soup = BeautifulSoup(res,"html.parser")
            book_div = soup.find(attrs={"class":"r"})
            book_a = book_div.findAll(attrs={"rel":"chapter"})
            book_span = book_div.findAll(attrs={"class":"aux light_narrow_text"})
            book_link = book_div.findAll(attrs={"class":"avail"})

            for k in range(len(book_a)):            
                link = 'https://www.ncbi.nlm.nih.gov/m/pubmed' + book_a[k]['href'].replace('.', '')
                paperInfo = book_span[k].string.replace('\n', '').split('.')
                sheet.write(paperCount + 1, 1, paperInfo[0]) #author
                sheet.write(paperCount + 1, 2, paperInfo[1].strip()) #Jour
                sheet.write(paperCount + 1, 3, paperInfo[2]) #time
                book_a[k].string = book_a[k].string.replace('"', '').replace(u'\xa0', u'')
                if len(book_a[k].string) > 255:
                    book_a[k].string = book_a[k].string[0:250]
                mylink = 'HYPERLINK("%s";"%s")' % (link, book_a[k].string)
                sheet.write(paperCount + 1, 0, xlwt.Formula(mylink))
                paperCount += 1
            res.close()
            #if paperCount % 100 == 0:
            #   time.sleep(2)
        except IOError,x:
            print x
            i = i - 1
            continue

xlsName = 'D:\\ncbi-paper\\' + searchKwd + '.xls'
book.save(xlsName)
print '\r\nSearch done! Totally got %d papers of %s, please check %s' %(paperCount, searchKwd, xlsName)





