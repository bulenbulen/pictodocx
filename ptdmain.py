import picadd
import xlsadd


sheetp = '电脑端（PC端）'
filepath = r'C:/Users/LENOVO/Desktop/jt/test.docx'
listpath= r'C:/Users/LENOVO/Desktop/jt/功能模块清单.xlsx'
picfiles = r'C:/Users/LENOVO/Desktop/jt/all'
#modname函数返回最后一个非空值，modpathname函数返回xx-xx-xx-xx格式的字符串列表
modlist = xlsadd.modname(listpath, sheetp)
modpathnamelist = xlsadd.modpathname(listpath, sheetp)
#print(modlist)
for i in modlist:
    #print(i)
    try:
        #modpathnamelist对应modlist的最后一个非空值，作为图片标题；也可直接用modlist的值作为图片标题
        j=modpathnamelist[modlist.index(i)]
        #picaddname(docfile,picfile,picname,rownum=n,modpathname=j)可选
        picadd.picaddname(filepath, picfiles, i, rownum=0, modpathname=j  )
        print(i)
    except:
        pass


