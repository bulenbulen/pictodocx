import picadd
import xlsadd


filepath = r'D:/DATA/doctest/test.docx'
listpath= r'D:/DATA/doctest/list.xlsx'
picfiles = r'D:/DATA/doctest'
#modname函数返回最后一个非空值，modpathname函数返回xx-xx-xx-xx格式的字符串列表
modlist = xlsadd.modname(listpath)
modpathnamelist = xlsadd.modpathname(listpath)
#print(modlist)
for i in modlist:
    try:
        #modpathnamelist对应modlist的最后一个非空值，作为图片标题；也可直接用modlist的值作为图片标题
        j=modpathnamelist[modlist.index(i)]
        #picaddname(docfile,picfile,picname,rownum=n,modpathname=j)可选
        picadd.picaddname(filepath, picfiles, i)
        print(i)
    except:
        pass


