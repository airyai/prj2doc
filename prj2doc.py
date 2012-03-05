#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Project to Document utility.
Written by PWX <airyai@gmail.com> in Feb. 2012.

This utility will collect project sources under current 
directory, highlight their syntax, and output as a whole 
document.
'''

from __future__ import unicode_literals, print_function
import os, sys, fnmatch
import re, getopt
import codecs
#sys.setdefaultencoding('utf-8')

import chardet
from pygments import highlight
from pygments.lexers import guess_lexer_for_filename
from pygments.formatters import get_formatter_for_filename, HtmlFormatter
from pygments.styles import get_style_by_name, get_all_styles
from pygments.util import ClassNotFound

# templates
ALLOW_WHITESPACE_FILTER = set(('.c', '.cpp', '.cc', '.h', '.hpp', '.cs',
                              '.vb', '.js', '.php', '.java'))
DEFAULT_PATTERNS = ('makefile*', '*.cpp', '*.cc', '*.hpp', '*.c', '*.h',
                    '*.py', '*.pyw', '*.java', '*.cs', '*.vb', '*.js',
                    '*.php*')

DEFAULT_COMMENT_START = '//'
COMMENT_STARTS = {}
COMMENT_STARTS['Python'] = '#'
COMMENT_STARTS['Makefile'] = '#' 
COMMENT_STARTS['VB.net'] = '\''

FILE_HEADER_TEMPLATE = '''
文件：{filename}
语法：{language}
行数：{line}
'''.strip()

CURRENT_DIR_PREFIX = '.' + os.path.sep
def generate_header(filename, language, line):
    if (filename.startswith(CURRENT_DIR_PREFIX)):
        filename = filename[len(CURRENT_DIR_PREFIX):]
    ret = FILE_HEADER_TEMPLATE.format(filename=filename, language=language,
                                      line=line).split('\n')
    width = 0
    for l in ret:
        lw = 0
        for c in l:
            if (ord(c) < 256):
                lw += 1
            else:
                lw += 2
        width = max(width, lw)
    cmstart = COMMENT_STARTS.get(language, DEFAULT_COMMENT_START)
    hr = '{0}{1}'.format(cmstart, '=' * (width+1))
    
    ret = [hr] + ['{0} {1}'.format(cmstart, l) for l in ret] + [hr]
    return '\n'.join(ret)
    

# utilities
def readfile(path):
    with open(path, 'rb') as f:
        return f.read()
def writefile(path, cnt):
    with open(path, 'wb') as f:
        f.write(codecs.BOM_UTF8)
        f.write(cnt.encode('utf-8'))

# parse arguments
SHORT_OPT_PATTERN = 'ho:s:m:l:'
LONG_OPT_PATTERN = ('output=', 'style=', 'makefile=', 'help', 'list-style',
                    'linenos=')
def usage():
    print ('Usage: prj2doc [选项] ... [输入通配符] ...')
    print ('Written by 平芜泫 <airyai@gmail.com>。')
    print ('')
    print ('如果没有指定输入通配符，那么当前目录、以及子目录下的所有文件将被选择。\n'
           '否则将在当前目录和子目录下搜索所有符合通配符的文件。如果所有被选择的\n'
           '文件中存在 Makefile，那么文件的排列顺序将参照 Makefile 中第一次出现的\n'
           '顺序。')
    print ('')
    print ('  -o, --output=         输出文档的路径。文档类型会根据扩展名猜测。')
    print ('                        如果没有指定，则输出 project.html 和 project.doc。')
    print ('                        支持的扩展名：html, doc, tex。')
    print ('                        注：doc 只在 Windows 下可用，且转换后需要手工微调。')
    print ('  -m, --makefile=       指定一个 Makefile 文件。')
    print ('  -s, --style=          设定代码高亮的配色方案。默认为 colorful。')
    print ('      --list-style           列出所有支持的配色方案。')
    print ('  -l, --lineno=[on/off] 打开或关闭源文件每一行的行号。')
    print ('  -h, --help            显示这个信息。')
    
def listStyles():
    print (' '.join(get_all_styles()))
    
optlist, args = getopt.getopt(sys.argv[1:], SHORT_OPT_PATTERN, LONG_OPT_PATTERN)
OUTPUT = []
MKPATTERN = '*makefile*'
MAKEFILE = None
STYLE = 'colorful'
LINENOS = True
def GET_LINENOS(fmt):
    return True if LINENOS else False

for (k, v) in optlist:
    if (k in ('-o', '--output=')):
        OUTPUT.append(v)
    elif (k in ('-s', '--style=')):
        STYLE = v
    elif (k in ('--list-style',)):
        listStyles()
        sys.exit(0)
    elif (k in ('-l', '--lineno=')):
        LINENOS = (v == 'on')
    elif (k in ('-m', '--makefile=')):
        MKPATTERN = v
    elif (k in ('-h', '--help')):
        usage()
        sys.exit(0)
        
PATTERNS = args
if (len(PATTERNS) == 0):
    PATTERNS = DEFAULT_PATTERNS
if (len(OUTPUT) == 0):
    OUTPUT = ['project.html']
if (sys.platform == 'win32'):
    OUTPUT.append('project.doc')

# scan input files
FORBIDS = ('prj2doc*', )
def scan_dir(path, file_list):
    global MAKEFILE
    dir_list = []
    # first list all files under the directory
    for p in os.listdir(path):
        p2 = os.path.join(path, p)
        if (os.path.isdir(p2)):
            dir_list.append(p2)
        else:
            flag = True
            for pattern in FORBIDS:
                if (fnmatch.fnmatchcase(p, pattern)):
                    flag = False
            if (not flag):
                continue
            for pattern in PATTERNS:
                if (fnmatch.fnmatchcase(p, pattern)):
                    file_list.append(os.path.normcase(p2))
            if (MAKEFILE is None and fnmatch.fnmatchcase(p, MKPATTERN)):
                MAKEFILE = os.path.normcase(p2)
    # then recursively scan
    for p2 in dir_list:
        scan_dir(p2, file_list)

INPUTS = []
print ('正在扫描目录下的所有源文件...')
scan_dir('.', INPUTS)

# check Makefile
phrase_map = {}
if (MAKEFILE is not None):
    makefile = readfile(MAKEFILE)
    regex = re.compile('\\s+')
    phrases = regex.split(makefile.lower())
    phrases.insert(0, MAKEFILE)
    phindex = 0
    for p in phrases:
        p = os.path.normcase(p)
        if (os.path.isfile(p)):
            n = os.path.split(p)[1]
            n = os.path.splitext(n)[0]
            phrase_map.setdefault(n, phindex)
            phindex += 1

def ext_compare(x, y):
    if (x == '.h' and y == '.cpp'):
        return -1
    elif (x == '.cpp' and y == '.h'):
        return 1
    else:
        return cmp(x, y)

def filename_compare(x, y):
    xx = os.path.splitext(os.path.split(x)[1])
    yy = os.path.splitext(os.path.split(y)[1])
    xi = phrase_map.get(xx[0], None)
    yi = phrase_map.get(yy[0], None)
    if (xi is not None and yi is not None):
        ret = cmp(xi, yi)
        if (ret != 0):
            return ret
        return ext_compare(xx[1], yy[1])
    elif (xi is not None):
        return -1
    elif (yi is not None):
        return 1
    else:
        ret = cmp(xx[0], yy[0])
        if (ret != 0):
            return ret
        return ext_compare(xx[1], yy[1])

INPUTS.sort(filename_compare)

# convert via MS Office
try:
    import win32com.client
    WIN32_SUPPORT = True
except ImportError:
    WIN32_SUPPORT = False
    pass

CONV_TEMP = None
CONV_LIST = []
if WIN32_SUPPORT:
    # define convert function
    def html2doc(htmlPath, docPath):
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(os.path.abspath(htmlPath).encode(sys.getfilesystemencoding()))
        doc.SaveAs(os.path.abspath(docPath).encode(sys.getfilesystemencoding()), FileFormat=0)
        doc.Close()
        word.Quit()
    # process list
    CONV_LIST = []
    new_output = []
    CONV_TEMP = 'prj2doc.temp.html'
    for i in range(0, len(OUTPUT)):
        if (os.path.splitext(OUTPUT[i])[1].lower() in ('.doc', )):
            CONV_LIST.append(OUTPUT[i])
        else:
            new_output.append(OUTPUT[i])
    if (len(CONV_LIST) > 0):
        new_output.append(CONV_TEMP)
        OUTPUT = new_output


# create formatters & lexers for output files
FORMATTERS = {}
CONTENTS = {}
LEXERS = {}

print ('读取源文件，并载入代码高亮引擎...')
try:
    STYLE = get_style_by_name(STYLE)
except ClassNotFound:
    print ('未定义的配色方案 {0}。'.format(STYLE))
    sys.exit(10)

for o in OUTPUT:
    try:
        f = get_formatter_for_filename(o)
        f.style = STYLE
        f.encoding = 'utf-8'
        #f.noclasses = True
        #f.nobackground = True
        FORMATTERS[o] = f
    except ClassNotFound:
        print ('不支持的输出格式 {0}。'.format(o))
        sys.exit(12)

def front_tab_to_space(x):
    for i in range(0, len(x)):
        if (x[i] != '\t'):
            return '    ' * i + x[i:]
    return '    ' * len(x)

for i in INPUTS:
    try:
        cnt = readfile(i)
        if (len(cnt) == 0):
            continue
        CONTENTS[i] = cnt
    except Exception as ex:
        print ('无法读取源文件 {0}：{1}。'.format(i, ex))
        #sys.exit(11)
    try:
        l = guess_lexer_for_filename(i, readfile(i))
        l.encoding = 'utf-8'
        LEXERS[i] = l
    except ClassNotFound:
        print ('不能确定源文件 {0} 的语法类型。'.format(i))
        #sys.exit(13)
        
if (len(CONTENTS) == 0):
    print ('没有找到任何非空输入文件。')
    sys.exit(0)

# generating sources for each file
CHARDET_REPLACE = {'gb2312': 'gb18030', 'gbk': 'gb18030'}
def detect_encoding(cnt):
    ret = chardet.detect(cnt)['encoding']
    if (ret is None):
        ret = sys.getfilesystemencoding()
    ret = ret.lower()
    return CHARDET_REPLACE.get(ret, ret)

HIGHLIGHTS = {o:[] for o in OUTPUT}
HIGHLIGHT_STYLES = {}
for k in INPUTS:
    if (k not in CONTENTS or k not in LEXERS):
        continue
    print ('正在处理 {0} ...'.format(k))
    lexer = LEXERS[k]
    cnt = CONTENTS[k]
    encoding = detect_encoding(cnt)
    if (encoding != 'gb18030' and encoding != 'utf-8'):
        encoding = sys.getfilesystemencoding() # Special hack!!!
    cnt = unicode(cnt, encoding)
    if (os.path.splitext(k)[1].lower() in ALLOW_WHITESPACE_FILTER):
        cnt = '\n'.join([front_tab_to_space(x)
                        for x in cnt.split('\n')])
    for o in OUTPUT:
        if (o not in FORMATTERS):
            continue
        f = FORMATTERS[o]
        header = generate_header(k, lexer.name, cnt.count('\n') + 1)
        # header
        f.linenos = False
        HIGHLIGHTS[o].append(unicode(highlight(header, lexer, f), 'utf-8'))
        # body
        f.linenos = GET_LINENOS(f) if (o != CONV_TEMP) else False
        f.nobackground = (o == CONV_TEMP)
        HIGHLIGHTS[o].append(unicode(highlight(cnt, lexer, f), 'utf-8'))
        # style
        if (o not in HIGHLIGHT_STYLES and hasattr(f, 'get_style_defs')):
            HIGHLIGHT_STYLES[o] = '\n'.join([f.get_style_defs('')])

# combining outputs
COMBINE_TEMPLATE_HTML = '''
<html>
    <head>
        <title>Project Document</title>
        <meta http-equiv="Content-Type" content="text/html;charset=utf-8;" />
        <style>
        pre {{ margin: 3px 2px; }}
        .linenodiv {{
            background: #eeeeee;
            padding-right: 1px;
            margin-right: 2px;
            text-align: right;
        }}
        * {{
            font-size: 13px;
            font-family: WenQuanYi Micro Hei Mono, 微软雅黑, Droid Sans, DejaVu Sans Mono, monospace;
        }}
        {style}
        </style>
    </head>
    <body>
        {body}
    </body>
</html>
'''.strip()
def combine_html(outputs, path):
    ret = []
    for i in range(0, len(outputs), 2):
        ret.append('<p>{0}\n{1}</p>'.format(outputs[i], outputs[i+1]))
    return COMBINE_TEMPLATE_HTML.format(body='\n<p>&nbsp;</p>\n'.join(ret),
                                        style=HIGHLIGHT_STYLES.get(path, ''))

def combine_other(outputs, path):
    return '\n'.join(outputs)

COMBINE_TABLE = {'.html': combine_html, '.htm': combine_html}
print ('将结果写入指定的输出 ...')
for o in OUTPUT:
    try:
        writefile(o, COMBINE_TABLE.get(os.path.splitext(o)[1].lower(),
                                       combine_other)(HIGHLIGHTS[o], o))
    except Exception as ex:
        print ('写入文件 {0} 失败：{1}。'.format(o, ex))

# do office convert
if (WIN32_SUPPORT and os.path.isfile(CONV_TEMP)):
    conv_body = unicode(readfile(CONV_TEMP), 'utf-8')
    for cv in CONV_LIST:
        try:
            html2doc(CONV_TEMP, cv)
        except Exception as ex:
            writefile(os.path.join(os.path.dirname(CONV_TEMP), cv + ".html"), conv_body)
            print ('转换文件为 {0} 失败，保留中间文件 {0}.html。'.format(cv))
    os.remove(CONV_TEMP)



