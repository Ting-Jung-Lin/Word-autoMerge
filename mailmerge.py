from copy import deepcopy
import warnings
from lxml.etree import Element
from lxml import etree
from zipfile import ZipFile, ZIP_DEFLATED
import shlex
import cn2an
from tkinter import messagebox

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)

CONTENT_TYPE_SETTINGS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'


class MailMerge(object):
    def __init__(self, file, remove_empty_tables=False):
        self.zip = ZipFile(file)
        self.parts = {}
        self.settings = None
        self._settings_info = None 
        self.remove_empty_tables = remove_empty_tables
        try:
            content_types = etree.parse(self.zip.open('[Content_Types].xml'))
            for file in content_types.findall('{%(ct)s}Override' % NAMESPACES):
                type = file.attrib['ContentType' % NAMESPACES]
                if type in CONTENT_TYPES_PARTS:
                    zi, self.parts[zi] = self.__get_tree_of_file(file)
                elif type == CONTENT_TYPE_SETTINGS:
                    self._settings_info, self.settings = self.__get_tree_of_file(file)

            to_delete = []
            for part in self.parts.values():
                for parent in part.findall('.//{%(w)s}fldSimple/..' % NAMESPACES):
                    for idx, child in enumerate(parent):
                        if child.tag != '{%(w)s}fldSimple' % NAMESPACES:
                            continue
                        instr = child.attrib['{%(w)s}instr' % NAMESPACES]
                        name = self.__parse_instr(instr)
                        if name is None:
                            continue
                        parent[idx] = Element('MergeField', name=name)

                for parent in part.findall('.//{%(w)s}instrText/../..' % NAMESPACES):#p
                    children = list(parent)
                    fields = zip(
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..' % NAMESPACES)],#r
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="end"]/..' % NAMESPACES)]
                    )
                    for idx_begin, idx_end in fields:
                        # consolidate all instrText nodes between'begin' and 'end' into a single node
                        begin = children[idx_begin]
                        instr_elements = [e for e in
                                          begin.getparent().findall('{%(w)s}r/{%(w)s}instrText' % NAMESPACES)#p
                                          if idx_begin < children.index(e.getparent()) < idx_end]
                        if len(instr_elements) == 0:
                            continue
                        # set the text of the first instrText element to the concatenation
                        # of all the instrText element texts
                        instr_text = ''.join([e.text for e in instr_elements])
                        instr_elements[0].text = instr_text

                        # delete all instrText elements except the first
                        for instr in instr_elements[1:]:
                            instr.getparent().remove(instr) #r #只剩instr_elements[0]
                        name = self.__parse_instr(instr_text)
                        if name is None:
                            continue
                        parent[idx_begin] = Element('MergeField', name=name)#<w:r --> <MergeField
                        # use this so we know *where* to put the replacement
                        #print(etree.tostring(instr_elements[0]))#<w:instrText
                        instr_elements[0].tag = 'MergeText'
                        block = instr_elements[0].getparent()#r#裏頭只有instr_elements[0]
                        # append the other tags in the w:r block too
                        parent[idx_begin].extend(list(block))#idx_begin 是 r 的標籤
                        #print(etree.tostring(parent))
                        to_delete += [(parent, parent[i + 1])
                                      for i in range(idx_begin, idx_end)]#剩parent[begin]跟之前的沒有刪掉
            for parent, child in to_delete:
                parent.remove(child)
            # Remove mail merge settings to avoid error messages when opening document in Winword
            if self.settings:
                settings_root = self.settings.getroot()
                mail_merge = settings_root.find('{%(w)s}mailMerge' % NAMESPACES)
                if mail_merge is not None:
                    settings_root.remove(mail_merge)
            
        except:
            self.zip.close()
            raise

    @classmethod
    def __parse_instr(cls, instr):
        args = shlex.split(instr, posix=False)
        if args[0] != 'MERGEFIELD':
            return None
        name = args[1]
        if name[0] == '"' and name[-1] == '"':
            name = name[1:-1]
        return name

    def __get_tree_of_file(self, file):
        fn = file.attrib['PartName' % NAMESPACES].split('/', 1)[1]
        zi = self.zip.getinfo(fn) #  = ZipFile(file).getinfo(word/document.xml)
        #不知道zip.open參數可不可以傳進壓縮資料夾？
        #print(zi)#<ZipInfo filename='word/document.xml' compress_type=deflate file_size=41219 compress_size=2102>
        return zi, etree.parse(self.zip.open(zi))
#輸出檔案
    def write(self, file):
        # Replace all remaining merge fields with "empty" values
        for field in self.get_merge_fields():
            self.merge(**{field: ''})
        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zi in self.zip.filelist:
                if zi in self.parts: #CONTENT_TYPES_PARTS
                    xml = etree.tostring(self.parts[zi].getroot())
                    output.writestr(zi.filename, xml)
                elif zi == self._settings_info:
                    xml = etree.tostring(self.settings.getroot())
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.zip.read(zi))
    def get_merge_fields(self, parts=None):
        if not parts:  #if parts==None
            parts = self.parts.values()
        else:
            parts = parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall('.//MergeField'):
                fields.add(mf.attrib['name'])
        return fields #標籤名字set

    def merge_templates(self, replacements, separator):
        """
        Duplicate template. Creates a copy of the template, does a merge, and separates them by a new paragraph, a new break or a new section break.
        separator must be :
        - page_break : Page Break. 
        - column_break : Column Break. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - textWrapping_break : Line Break.
        - continuous_section : Continuous section break. Begins the section on the next paragraph.
        - evenPage_section : evenPage section break. section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        - nextColumn_section : nextColumn section break. section begins on the following column on the page. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - nextPage_section : nextPage section break. section begins on the following page.
        - oddPage_section : oddPage section break. section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        """

        #TYPE PARAM CONTROL AND SPLIT
        valid_separators = {'page_break', 'column_break', 'textWrapping_break', 'continuous_section', 'evenPage_section', 'nextColumn_section', 'nextPage_section', 'oddPage_section'}
        if not separator in valid_separators:
            raise ValueError("Invalid separator argument")
        type, sepClass = separator.split("_")
  

        #GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue
		
            if sepClass == 'section':

                #FINDING FIRST SECTION OF THE DOCUMENT
                firstSection = root.find("w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES)
                if firstSection == None:
                    firstSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)
			
                #MODIFY TYPE ATTRIBUTE OF FIRST SECTION FOR MERGING
                nextPageSec = deepcopy(firstSection)
                for child in nextPageSec:
                #Delete old type if exist
                    if child.tag == '{%(w)s}type' % NAMESPACES:
                        nextPageSec.remove(child)
                #Create new type (def parameter)
                newType = etree.SubElement(nextPageSec, '{%(w)s}type'  % NAMESPACES)
                newType.set('{%(w)s}val'  % NAMESPACES, type)

                #REPLACING FIRST SECTION
                secRoot = firstSection.getparent()
                secRoot.replace(firstSection, nextPageSec)

            #FINDING LAST SECTION OF THE DOCUMENT
            lastSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)

            #SAVING LAST SECTION
            mainSection = deepcopy(lastSection)
            lsecRoot = lastSection.getparent()
            lsecRoot.remove(lastSection)

            #COPY CHILDREN ELEMENTS OF BODY IN A LIST
            childrenList = root.findall('w:body/*', namespaces=NAMESPACES)

            #DELETE ALL CHILDREN OF BODY
            for child in root:
                if child.tag == '{%(w)s}body' % NAMESPACES:
                    child.clear()

            #REFILL BODY AND MERGE DOCS - ADD LAST SECTION ENCAPSULATED OR NOT
            lr = len(replacements)
            lc = len(childrenList)

            for i, repl in enumerate(replacements):
                parts = []
                for (j, n) in enumerate(childrenList):
                    element = deepcopy(n)
                    for child in root:
                        if child.tag == '{%(w)s}body' % NAMESPACES:
                            child.append(element)
                            parts.append(element)
                            if (j + 1) == lc:
                                if (i + 1) == lr:
                                    child.append(mainSection)
                                    parts.append(mainSection)
                                else:
                                    if sepClass == 'section':
                                        intSection = deepcopy(mainSection)
                                        p   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        pPr = etree.SubElement(p, '{%(w)s}pPr'  % NAMESPACES)
                                        pPr.append(intSection)
                                        parts.append(p)
                                    elif sepClass == 'break':
                                        pb   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        r = etree.SubElement(pb, '{%(w)s}r'  % NAMESPACES)
                                        nbreak = Element('{%(w)s}br' % NAMESPACES)
                                        nbreak.attrib['{%(w)s}type' % NAMESPACES] = type
                                        r.append(nbreak)

                    self.merge(parts, **repl)

    def merge_pages(self, replacements):
         """
         Deprecated method.
         """
         warnings.warn("merge_pages has been deprecated in favour of merge_templates",
                      category=DeprecationWarning,
                      stacklevel=2)         
         self.merge_templates(replacements, "page_break")

   
                
#合併種類判斷
    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()
        else:
            parts = parts.values()
        for field, replacement in replacements.items():
            if isinstance(replacement, list):
                self.merge_rows(field, replacement)
            else:
                for part in parts:  
                    self.__merge_field(part, field, replacement)
#實際合併
    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            children = list(mf)
            mf.clear()  # clear away the attributes
                #print(etree.tostring(mf))#扣掉name屬性
            mf.tag = '{%(w)s}r' % NAMESPACES
                #print(etree.tostring(mf))#把 Mergefield 改成 w:r
            mf.extend(children)
                #print(etree.tostring(mf))
            nodes = []
                # preserve new lines in replacement text
            text = text or ''  # text might be None

            text_parts = str(text).replace('\r', '').split('\n')# return list 
            
            for i, text_part in enumerate(text_parts):
                text_node = Element('{%(w)s}t' % NAMESPACES)
                text_node.text = text_part
                nodes.append(text_node)
                # if not last node add new line node
                if i < (len(text_parts) - 1):
                    nodes.append(Element('{%(w)s}br' % NAMESPACES))
            ph = mf.find('MergeText')
            if ph is not None:
                # add text nodes at the exact position where
                # MergeText was found
                index = mf.index(ph)
                for node in reversed(nodes): #為何要reversed
                #for node in nodes: 
                    mf.insert(index, node)
                mf.remove(ph)
            else:
                mf.extend(nodes)

    def merge_rows(self, anchor, rows):
        store=self.__find_row_anchor(anchor)
        for i in range(len(store)):
            table, idx, template =store[i]
        #table, idx, template = self.__find_row_anchor(anchor)
            if table is not None:
                if len(rows) > 0:
                    del table[idx]
                    row={}
                    for i, row_data in enumerate(rows):
                        row[i] = deepcopy(template)
                        self.merge(row, **row_data)
                        table.insert(idx + i, row[i])
                else:
                    # if there is no data for a given table
                    # we check whether table needs to be removed
                    if self.remove_empty_tables:
                        parent = table.getparent()
                        parent.remove(table)

    def __find_row_anchor(self, field, parts=None):
        store=[]
        if not parts:
            parts = self.parts.values()
        
        for part in parts:
            for table in part.findall('.//{%(w)s}tbl' % NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        store.append((table,idx,row))
        return store
                        #return table, idx, row
    
    def field_concate(self, anchor, rows, symbol,parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for mf in part.findall('.//MergeField[@name="%s"]' % anchor):
                double = deepcopy(mf)
                for i in range(len(rows)):
                    self.__merge_field(mf.getparent(), anchor, rows[i][anchor])
                    if( i != len(rows)-1):
                        dash=Element('{%(w)s}r' % NAMESPACES)
                        attri=deepcopy(mf.getparent().find('.//{%(w)s}rPr' % NAMESPACES))
                        dash.append(attri)
                        dash2=etree.SubElement(dash,'{%(w)s}t' % NAMESPACES)
                        dash2.text=symbol
                        mf.getparent().insert((mf.getparent().index(mf)+1+2*i),dash)
                        double2=deepcopy(double)
                        mf.getparent().insert((mf.getparent().index(mf)+2+2*i),double2)
    '''
    def p_concate2(self,start,end,rows,anchor,symbol):
        parent=etree.Element('parent')
        for word in range(comma_index,end_index+1):
            temp=deepcopy(bookmark.getparent().getparent()[index+4*i+paragraph][word])
            parent.append(temp)
        double=deepcopy(parent)
    '''
    def p_concate(self,anchor, rows,symbol):
        parts = self.parts.values()
        for part in parts:
            root = part.getroot()
            flag=part.findall('.//MergeField[@name="%s"]' % anchor)
            for i in range(len(flag)):
                double = deepcopy(list(flag[i].getparent()))
                parent_end_index=len(flag[i].getparent())-1
                for j in range(len(rows)):
                    temp_parent=etree.Element('parent')
                    double2=deepcopy(double)
                    for child in double2:
                        temp_parent.append(child)
                    for mfn in flag[i].getparent().findall('.//MergeField'):
                        self.__merge_field(flag[i].getparent(), mfn.get("name"), rows[j][mfn.get("name")])
                    if( j != len(rows)-1):
                        dash=Element('{%(w)s}r' % NAMESPACES)
                        dash2=etree.SubElement(dash,'{%(w)s}t' % NAMESPACES)
                        dash2.text=symbol
                        flag[i].getparent().append(dash)
                        for child in temp_parent:
                            flag[i].getparent().append(child)
    
    
    def for_short_multi(self,diff_period_list,period):
        parts = self.parts.values()
        
        for part in parts:
            bookmarks=part.findall('.//{%(w)s}bookmarkStart' % NAMESPACES)
            for bookmark in bookmarks:
                if(bookmark.get("{%(w)s}name"% NAMESPACES)=="短路容量"):
                    short_list=[]
                    #"一期"段落
                    parent=etree.Element('parent')
                    index=bookmark.getparent().getparent().index(bookmark.getparent())
                    for i in range(4):
                        parent.append(deepcopy(bookmark.getparent().getparent()[index+i]))
                    #開始複製
                    for i in range(period):
                        #不是一期，複製，若是0(也就是一期)，不執行
                        if(i!=0):
                            #複製
                            double=deepcopy(parent)
                            for double_mark_start in double.findall('.//{%(w)s}bookmarkStart' % NAMESPACES):
                                if(double_mark_start.get('{%(w)s}name' % NAMESPACES)=="短路容量"):
                                    for double_mark_end in double.findall('.//{%(w)s}bookmarkEnd' % NAMESPACES):
                                        if(double_mark_end.get('{%(w)s}id' % NAMESPACES)==double_mark_start.get('{%(w)s}id' % NAMESPACES)):
                                            double_mark_end.getparent().remove(double_mark_end)
                                            break
                                    double_mark_start.getparent().remove(double_mark_start)
                                    break
                            #插入複製的段落
                            double[0].find('.//{%(w)s}r' % NAMESPACES).find('.//{%(w)s}t' % NAMESPACES).text=cn2an.an2cn(i+1)+"期："
                            for j in range(len(list(double))):
                                bookmark.getparent().getparent().insert(index+4*i+j,double[0])
                        
                        
                        diff_inv=[]
                        for row in range(len(diff_period_list)):
                            if(cn2an.cn2an(diff_period_list[row]["期別"][0])==i+1):
                                diff_inv.append(diff_period_list[row])
                                
                        #找part
                        copy={}
                        #處理diff_inv
                        #填入資料(某一期)
                        if(len(diff_inv))>1:
                            for paragraph in range(4):
                                if(i+1==cn2an.cn2an(diff_inv[0]["期別"][0])):
                                    if(paragraph==1):
                                        
                                        comma_index=bookmark.getparent().getparent()[index+4*i+paragraph].index(bookmark.getparent().getparent()[index+4*i+paragraph].find('.//MergeField[@name="逆變器額定輸出功率千瓦"]'))-1
                                        end_index=len(bookmark.getparent().getparent()[index+4*i+paragraph])-1
                                        
                                        #self.p_concate2(comma_index,end_index,anchor,diff_inv,("+"))
                                        
                                        parent2=etree.Element('parent')
                                        for word in range(comma_index,end_index+1):
                                            temp=deepcopy(bookmark.getparent().getparent()[index+4*i+paragraph][word])
                                            parent2.append(temp)
                                        double=deepcopy(parent2)
                                        #先填第一個
                                        for field in self.get_merge_fields(parts={0:bookmark.getparent().getparent()[index+4*i+paragraph]}):
                                            self.__merge_field(bookmark.getparent().getparent()[index+4*i+paragraph],field,diff_inv[0][field])
                                            try:
                                                count_short=float(diff_inv[0]["短路容量"])
                                            except ValueError:
                                                messagebox.showwarning(title="提醒",message=diff_inv[0]["識別碼"]+diff_inv[0]["設置者名稱"]+"無短路容量資料")
                                                return
                                        #再貼上新的
                                        for inv in range(1,len(diff_inv)):
                                            double2=deepcopy(double)
                                            try:
                                                count_short+=float(diff_inv[inv]["短路容量"])
                                            except ValueError:
                                                messagebox.showwarning(title="提醒",message=diff_inv[inv]["識別碼"]+diff_inv[0]["設置者名稱"]+"無短路容量資料")
                                                return
                                            #要複製的   
                                            for k in range(len(list(double))):
                                                bookmark.getparent().getparent()[index+4*i+paragraph].insert((end_index+1)+k+(end_index+1-comma_index)*(inv-1),double2[0])
                                                for field in self.get_merge_fields(parts={0:bookmark.getparent().getparent()[index+4*i+paragraph]}):
                                                    self.__merge_field(bookmark.getparent().getparent()[index+4*i+paragraph],field,diff_inv[inv][field])
                                    elif(paragraph==3):
                                        bracket_index=bookmark.getparent().getparent()[index+4*i+paragraph].index(bookmark.getparent().getparent()[index+4*i+paragraph].find('.//MergeField[@name="逆變器搭配模組片數"]'))
                                        end_index=bookmark.getparent().getparent()[index+4*i+paragraph].index(bookmark.getparent().getparent()[index+4*i+paragraph].find('.//MergeField[@name="短路容量"]'))-3
                                        symbol=deepcopy(bookmark.getparent().getparent()[index+4*i+paragraph].find('.//MergeField[@name="逆變器搭配模組片數"]').getprevious())
                                        symbol.find('.//{%(w)s}t' % NAMESPACES).text="+"
                                        parent2=etree.Element('parent')
                                        for word in range(bracket_index,end_index+1):
                                            temp=deepcopy(bookmark.getparent().getparent()[index+4*i+paragraph][word])
                                            parent2.append(temp)
                                        double=deepcopy(parent2)
                                         #先填第一個
                                        for field in self.get_merge_fields(parts={0:bookmark.getparent().getparent()[index+4*i+paragraph]}):
                                            if(field=="短路容量"):
                                                self.__merge_field(bookmark.getparent().getparent()[index+4*i+paragraph],field,round(count_short,4))
                                            else:
                                                self.__merge_field(bookmark.getparent().getparent()[index+4*i+paragraph],field,diff_inv[0][field])
                                        #再貼上新的
                                        for inv in range(1,len(diff_inv)):
                                            double2=deepcopy(double)
                                            #要複製的   
                                            for k in range(len(list(double))):
                                                bookmark.getparent().getparent()[index+4*i+paragraph].insert((end_index+1)+k+(end_index+1-bracket_index)*(inv-1),double2[0])
                                                for field in self.get_merge_fields(parts={0:bookmark.getparent().getparent()[index+4*i+paragraph]}):
                                                    self.__merge_field(bookmark.getparent().getparent()[index+4*i+paragraph],field,diff_inv[inv][field])
                                        #加入標點符號
                                        
                                        for inv in range(1,len(diff_inv)):
                                            symbol2=deepcopy(symbol)
                                            bookmark.getparent().getparent()[index+4*i+paragraph].insert((end_index+1)+(end_index+1-bracket_index)*(inv-1)+(inv-1),symbol2)
                            short_list.append({"短路容量":round(count_short,4)})
                        else:
                            bracket=bookmark.getparent().getparent()[index+4*i+3].find('.//MergeField[@name="逆變器搭配模組片數"]').getprevious()
                            end=bookmark.getparent().getparent()[index+4*i+3].find('.//MergeField[@name="短路容量"]').getprevious().getprevious()
                            bookmark.getparent().getparent()[index+4*i+3].remove(bracket)
                            bookmark.getparent().getparent()[index+4*i+3].remove(end)
                            try:
                                short_list.append({"短路容量":round(float(diff_inv[0]["短路容量"]),4)})
                            except ValueError:
                                messagebox.showwarning("提醒",diff_inv[0]["識別碼"]+diff_inv[0]["設置者名稱"]+"無短路容量資料")
                            
                            for paragraph in range(4):
                                copy[0]=bookmark.getparent().getparent()[index+4*i+paragraph]
                                for field in self.get_merge_fields(parts=copy):
                                    self.__merge_field(copy[0], field, diff_inv[0][field])
                    break
            #總短路容量填入資料
            for bookmark in bookmarks:#可以改成用找MERGEFIELD
                if(bookmark.get("{%(w)s}name"% NAMESPACES)=="總短路容量"):
                    mf=bookmark.getparent().find('.//MergeField')
                    self.field_concate("短路容量",short_list, "kVA+",parts=[mf.getparent()])
                    short=0
                    for short_dict in short_list:
                        short=short+short_dict["短路容量"]
                    short=round(short,4)
                    mf.getparent()[-1].find('.//{%(w)s}t' % NAMESPACES).text="kVA="+str(short)+"kVA"
                    self.__merge_field(part.find('.//MergeField[@name="短路容量"]').getparent(),"短路容量",short)
                    break
    
    def remove_short_first(self):
        parts = self.parts.values()
        for part in parts: 
            bookmarks=part.findall('.//{%(w)s}bookmarkStart' % NAMESPACES)
            for bookmark in bookmarks:
                if(bookmark.get("{%(w)s}name"% NAMESPACES)=="短路容量"):
                    index=bookmark.getparent().getparent().index(bookmark.getparent())
                    bookmark.getparent().getparent().remove(bookmark.getparent().getparent()[index])
                if(bookmark.get("{%(w)s}name"% NAMESPACES)=="總短路容量"):
                    index=bookmark.getparent().getparent().index(bookmark.getparent())
                    bookmark.getparent().getparent().remove(bookmark.getparent().getparent()[index])
                    
    def choice(self,merge_fields):
        parts = self.parts.values()
        for part in parts:
            for mf in part.findall('.//MergeField[@name="%s"]' % "設置地址"):
                index=mf.getparent().index(mf)
                parent_len=len(mf.getparent())
                if(merge_fields["設置地址"]=="None"):
                    mf.getparent().remove(mf)
                ##還有一個狀況是兩個都不是None，要移除地號
                elif(merge_fields["設置地號"]!="None"):
                    for siblings in mf.itersiblings():
                        if siblings.tag=="MergeField" and siblings.get("name")=="設置地號":
                            mf.getparent().remove(siblings)
            for mf in part.findall('.//MergeField[@name="%s"]' % "設置地號"):
                index=mf.getparent().index(mf)
                parent_len=len(mf.getparent())
                if(merge_fields["設置地號"]=="None"):
                    mf.getparent().remove(mf)
                elif(merge_fields["設置地址"]!="None"):
                    for siblings in mf.itersiblings():
                        if siblings.tag=="MergeField" and siblings.get("name")=="設置地址":
                            mf.getparent().remove(mf)
            
    def judge(self,merge_fields):
        parts = self.parts.values()
        for part in parts:
            
            if(merge_fields["建號"]!="None"):
                #既有建物
                for mf in part.findall('.//MergeField[@name="%s"]' % "建號"):
                    mf.getparent().find("{%(w)s}r"% NAMESPACES).find("{%(w)s}t"% NAMESPACES).text="■"
                    for mf_delete in mf.getparent().getnext().findall('.//MergeField'):
                        mf.getparent().getnext().remove(mf_delete)
            elif(merge_fields["建築執照號碼"]!="None"):
                #新建物
                for mf in part.findall('.//MergeField[@name="%s"]' % "建築執照號碼"):
                    mf.getparent().find("{%(w)s}r"% NAMESPACES).find("{%(w)s}t"% NAMESPACES).text="■"
                    for mf_delete in mf.getparent().getprevious().findall('.//MergeField'):
                        mf.getparent().getprevious().remove(mf_delete)
    def remove_period(self):
        parts = self.parts.values()
        for part in parts:
            for mf in part.findall('.//MergeField[@name="期別"]'):
                if(mf!=None):
                    mf.getparent().remove(mf)
    def count_sell(self,diff_period_list,period):
        parts = self.parts.values()
        for part in parts:
            existing_capacity=0
            total_capacity=0
            for i in range(period):
                for j in range(len(diff_period_list)):
                    if(cn2an.cn2an(diff_period_list[j]["期別"][0])==i+1):
                        if(cn2an.cn2an(diff_period_list[j]["期別"][0])<period):
                            existing_capacity+=float(diff_period_list[j]["躉售容量"])
                        total_capacity+=float(diff_period_list[j]["躉售容量"])
                        break
            for no in range(len(part.findall('.//MergeField[@name="躉售容量"]'))-1,-1,-1):
                if(no==5):
                    self.__merge_field(part.findall('.//MergeField[@name="躉售容量"]')[no].getparent(), "躉售容量", str(round(total_capacity,4)))
                elif(no==4):
                    self.__merge_field(part.findall('.//MergeField[@name="躉售容量"]')[no].getparent(), "躉售容量", str(round(total_capacity,4)))
                elif(no==1):
                    self.__merge_field(part.findall('.//MergeField[@name="躉售容量"]')[no].getparent(), "躉售容量", str(existing_capacity))
                elif(no==0):
                    self.__merge_field(part.findall('.//MergeField[@name="躉售容量"]')[no].getparent(), "躉售容量", str(existing_capacity))

            break
                        
                        
    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def close(self):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None