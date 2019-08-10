#!/usr/bin/python3
import re
import os
import sys
import glob
import xlsxwriter

from configparser import ConfigParser, ParsingError,ExtendedInterpolation
from collections import OrderedDict

term_regex = r'''(?mx)
    \s*(?:
        (?P<brackl>\()|
        (?P<brackr>\))|
        (?P<sq>"[^"]*")|
        (?P<s>[^(^)\s]+)
       )'''

def parse_sexp(sexp):
    stack = []
    out = []
    for termtypes in re.finditer(term_regex, sexp):
        term, value = [(t,v) for t,v in termtypes.groupdict().items() if v][0]
        if   term == 'brackl':
            stack.append(out)
            out = []
        elif term == 'brackr':
            assert stack, "Trouble with nesting of brackets"
            tmpout, out = out, stack.pop(-1)
            out.append(tmpout)
        elif term == 'sq':
            out.append(value[1:-1])
        elif term == 's':
            out.append(value)
        else:
            raise NotImplementedError("Error: %r" % term)
    assert not stack, "Trouble with nesting of brackets"
    return out[0]

def print_sexp(exp):
    out = ''
    if type(exp) == type([]):
        out += '(' + ' '.join(print_sexp(x) for x in exp) + ')'
    elif type(exp) == type('') and re.search(r'[\s()]', exp):
        out += '"%s"' % repr(exp)[1:-1].replace('"', '\"')
    else:
        out += '%s' % exp
    return out

# helper functions 

def sortRef(lst):
    return sorted(lst,key = lambda x: int('0'+''.join(re.findall(r'\d+', x))))

def listtonumbers(l):
    res = []
    for i in l:
        res.append(float(i))
    return res

# class definitions

class Module:
    def __init__(self,mod):
        self.module = mod
        self.used = False
    
    def getAttr(self,attribute):
        if attribute == "reference":
            return self.getRef()
        elif attribute == "package":
            return self.getPackage()
        elif attribute == "value":
            return self.getValue()
        elif attribute == "library":
            return self.getLib()
        elif attribute == 'coord':
            return ",".join(self.getCoord())    
        elif attribute == "angle":
            return self.getAngle()
        elif attribute == "side":
            return self.getSide()    
        elif attribute == 'category':
            return self.elementCategory([])
            
    def getRef(self):
        if not hasattr(self,'ref'):
            self.ref = ""
            for i in self.module:
                if isinstance(i,list) and i[0] == 'fp_text' and i[1] == 'reference':
                    self.ref = i[2]
        return self.ref
        
    def getPackage(self):
        if not hasattr(self,'package'):
            module = self.module[1].split(':')
            if len(module) == 1:
                self.package = module[0]
            else:
                self.package = module[1]
        return self.package
        
    def isSMD(self):
        if not hasattr(self,'smd'):
            self.smd = False
            for i in self.module:
                if isinstance(i,list) and len(i) > 1 and i[0] == 'attr' and i[1] == 'smd':
                    self.smd = True
        return self.smd
    
    def getLayer(self):
        if not hasattr(self,'layer'):
            self.layer = ''
            for i in self.module:
                if isinstance(i,list) and len(i) > 1 and i[0] == 'layer':
                    self.layer = i[1]
        return self.layer
    
    def getSide(self):
        layer = self.getLayer()
        if layer[:1] == 'F':
            return 'top'
        if layer[:1] == 'B':
            return 'bottom'
            
    def getAngle(self):
        c = self.getCoord()
        if len(c) == 3:
            return c[2]
        return 0
                
    def getValue(self):
        if not hasattr(self,'val'):
            self.val = ""
            for i in self.module:
                if isinstance(i,list) and i[0] == 'fp_text' and i[1] == 'value':
                    self.val = i[2]
        return self.val

    def getCoord(self):
        for i in self.module:
            if isinstance(i,list) and i[0] == 'at':
                return listtonumbers(i[1:])
        return []
    
    def getTags(self):
        for i in self.module:
            if isinstance(i,list) and i[0] == 'tags':
                res = i[1].split(',')
                return res
        return []        
       
    def getDescr(self):
        for i in self.module:
            if isinstance(i,list) and i[0] == 'descr':
                res = i[1].split(',')
                return res
        return []    
    
    def getLib(self):
        if not hasattr(self,'lib'):
            module = self.module[1].split(':')
            if len(module) == 1:
                self.lib = ''
            else:
                self.lib = module[0]
        return self.lib
    
    def getPads(self):
        res = []
        for i in self.module:
            if isinstance(i,list) and i[0] == 'pad':
                res.append(i)
        return res
    
    def padCoord(self,pad):
        for i in pad:
            if isinstance(i,list) and i[0] == 'at':
                return listtonumbers(i[1:])
        print("no coord for "+pad)
        return []
    
    def getCenter(self,origin=(0,0)):
        mod = self.getCoord()
        return (round(mod[0]-origin[0],2), abs(round(origin[1]-mod[1],2)))
    
    def isFiducial(self):
        if re.match(r"FID\d+",self.getRef()):
            return True
        if self.getLib().startswith("Fiducials"):
            return True
        for d in self.getDescr():
            if d.startswith("Fiducial"):
                return True
        return False
        
    def isResistor(self):
        for t in self.getTags():
            if t.startswith("resistor"):
                return True
        for d in self.getDescr():
            if d.startswith("Resistor"):
                return True
        if self.getLib().startswith("Resistors"):
            return True
        if re.match(r"R\d+",self.getRef()):
            return True
        return False
    
    def isCapacitor(self):
        for t in self.getTags():
            if t.startswith("capacitor"):
                return True
        for d in self.getDescr():
            if d.startswith("Capacitor"):
                return True
        if self.getLib().startswith("Capacitors"):
            return True
        if re.match(r"C\d+",self.getRef()):
            return True
        return False
    
    def isInductance(self):
        if re.match(r"L\d+",self.getRef()):
            return True
        return False
    
    def isTransistor(self):
        pc = len(self.getPads())
        if pc < 3:
            return False
        if re.match(r"Q\d+",self.getRef()):
            return True
        return False    
        
    def elementCategory(self,catlist):
        if not hasattr(self,'category'):
            for i in catlist:
                a = self.getAttr(i['attr'])
                if a and re.match(i['match'],str(a)):
                    self.category = i['category']
                    return self.category
            if self.isResistor():
                self.category = 'resistors'
            elif self.isCapacitor():
                self.category = 'capacitors'
            elif self.isInductance():
                self.category = 'inductances'
            elif self.isTransistor():
                self.category = 'transistors'
            else:    
                self.category = ""
        return self.category
            

class Board:
    def __init__(self,pname,options):
        self.options = options
        self.workbook = False
        with open(pname+".kicad_pcb", "r") as f:
            brd = f.read()
            self.brd = parse_sexp(brd)
            f.close()
        #with open(pname+".net", "r") as f:
        #    net = f.read()
        #    self.net = parse_sexp(net)
        #    f.close()
        self.modules = []
        for l in self.brd:
            if l[0] == 'module':
                m = Module(l)
                for p in self.options.package_sub:
                    if re.match(p['match'],m.getPackage()):
                        m.package = p['repl']
                self.modules.append(m)
                    
    def ignore(self,module):
        r = module.getRef()
        if r == '~' or r == '':
            return True
        for i in self.options.ignore:
            a = module.getAttr(i['attr'])
            if a and re.match(i['match'],a):
                print('Ignored',r)
                return True
        return False
       
    def listModules(self):
        origin = self.getPlaceOrigin()
        for l in self.modules:
            print(l.getRef(), l.getCenter(origin))
    
    def getPlaceOrigin(self):
        for l in self.brd:
            if isinstance(l,list) and l[0] == 'setup':
                for i in l:
                    if isinstance(i,list) and i[0] == 'aux_axis_origin':
                        return listtonumbers(i[1:])
        return []
    
    def prepareContents(self):
        sect = {}
        if self.hasSections():
            for s in self.options.sections:
                sect[s] = []
        else:
            sect[0] = []    
        for m in self.modules:
            if self.ignore(m):
                continue
            c = m.elementCategory(self.options.categories)
            if not self.hasSections():
                c = 0
            if c in sect:
                m.used = True
                section = sect[c]
                package = m.getPackage()
                value = m.getValue()
                found = False
                for i in section:
                    if i['package'] == package and i['value'] == value:
                        i['quantity'] += 1
                        i['reference'].append(m.getRef())
                        found = True
                if not found:
                    section.append(self.prepareModule(m))
        for m in self.modules:
            if not m.used and not self.ignore(m):
                if not '__default' in sect:
                    sect['__default'] = []
                sect['__default'].append(self.prepareModule(m))
        for s in sect:
            for i in sect[s]:
                i['reference'] = ','.join(i['reference'])
            if len(self.options.sections) ==0:
                sect[s] = sorted(sect[s],key = lambda x: x['reference'])
            else:
                sect[s] = sorted(sect[s],key = lambda x: x['key'])        
        self.contents = sect
        
    def prepareModule(self,module):
        modulerow = {'key': int('0'+''.join(re.findall(r'\d+', module.getRef()))), 'reference': [module.getRef()], 'package':module.getPackage(), 'value':module.getValue(),'quantity':1}
        for c in self.options.columns:
            attr = c['source']
            if attr != '' and not attr in modulerow:
                v = module.getAttr(attr)
                modulerow[attr] = v
        return modulerow
    
    def hasSections(self):
        return len(self.options.sections) > 0
    
    def createXLSX(self):
        if not self.workbook:
            filename = self.options.projectName+"_BOM.xlsx"
            self.workbook = xlsxwriter.Workbook(filename)
            for f in self.options.formats:
                self.options.formats[f] = self.workbook.add_format(self.options.formats[f])
        
    def writeXLSX(self):
        self.workbook.close() 
    
    def addBOM(self):
        worksheet = self.workbook.add_worksheet(self.options.projectName)
        # formats
        hdrfmt = self.options.formats['header']
        colhdrfmt = self.options.formats['column_header']
        nrfmt = self.options.formats['pos_number']
        sectfmt = self.options.formats['section_header']
        cellfmt = self.options.formats['cell']
        reffmt = self.options.formats['ref']
        # write data        
        worksheet.set_row(1, 25)
        worksheet.write(1, int(len(self.options.columns)/2), self.options.header,hdrfmt)
        worksheet.set_row(3, 50)
        i = 1
        for c in self.options.columns:
            worksheet.set_column(i,i, c['width'])
            worksheet.write(3,i,c['name'],colhdrfmt)
            i += 1
        self.prepareContents()
        row = 4
        n = 1
        sections = self.options.sections
        if not self.hasSections():
            sections = {0:''}
        for s in sections:
            if len(self.contents[s]) == 0: continue
            worksheet.set_row(row, 35)
            if s != 0:
                worksheet.write(row,3,sections[s],sectfmt)
                row += 1
            for m in self.contents[s]:
                worksheet.set_row(row, 25)
                col = 1
                for c in self.options.columns:
                    src = c['source']
                    if src:
                        if src == 'n':
                            worksheet.write(row,col,n,nrfmt)
                        elif src == 'reference':
                            worksheet.write(row,col,m[src],reffmt)    
                        else:
                            worksheet.write(row,col,m[src],cellfmt)
                    else:
                        worksheet.write(row,col,'',cellfmt)
                    col += 1
                row += 1
                n += 1
        if '__default' in self.contents:
            for m in self.contents['__default']:
                worksheet.write(row,1,str(m))  
                row += 1      
        
    
    def addPlacement(self):
        origin = self.getPlaceOrigin()
        worksheet = self.workbook.add_worksheet("component positions")
        # formats
        hdrfmt = self.options.formats['header']
        subhdrfmt = self.options.formats['subheader']
        colhdrfmt = self.options.formats['column_header']
        nrfmt = self.options.formats['pos_number']
        poshdr = self.options.config.get("project","pos_header",fallback = "Component positions")
        worksheet.set_row(0, 30)
        worksheet.write(0,1,poshdr,hdrfmt)
        row = 1
        try:
            worksheet.set_row(row, 25)
            worksheet.write(row,1,self.options.config.get("project","fid_header"),subhdrfmt)
            row += 1
        except:
            pass
        col_x = 3
        col_y = 4
        for c in range(0,len(self.options.pos_columns)):
            col = self.options.pos_columns[c]
            if col['source'] == 'x':
                col_x = c+1
            if col['source'] == 'y':
                col_y = c+1    
        worksheet.set_row(row, 25)
        worksheet.write(row,col_x,'X',colhdrfmt)
        worksheet.write(row,col_y,'Y',colhdrfmt)
        row += 1
        for m in self.modules:
            if m.isFiducial():
                coord = m.getCenter(origin)
                worksheet.write(row,col_x,coord[0])
                worksheet.write(row,col_y,coord[1])
                row += 1
        try:
            worksheet.set_row(row, 25)
            worksheet.write(row,1,self.options.config.get("project","pos_header"),subhdrfmt)
            row += 1
        except:
            pass
        col = 1    
        worksheet.set_row(row, 25)
        for c in self.options.pos_columns:
            worksheet.write(row,col,c['name'],colhdrfmt)
            worksheet.set_column(col,col,c['width'])
            col += 1
        row += 1
        n = 1
        for m in self.modules:
            if not m.isFiducial() and m.isSMD():
                coord = m.getCenter(origin)
                col = 1
                for c in self.options.pos_columns:
                    src = c['source']
                    if src == 'x':
                        worksheet.write(row,col,coord[0])
                    elif src == 'y':
                        worksheet.write(row,col,coord[1])
                    elif src == 'n':
                        worksheet.write(row,col,n,nrfmt)
                    else:
                        worksheet.write(row,col,m.getAttr(src))
                    col += 1
                n += 1
                row += 1
    
class Options:
    def __init__(self):
        self.config = ConfigParser(interpolation = ExtendedInterpolation(),allow_no_value=True)
        self.config.optionxform = lambda option: option
        try:
            self.config.read('bom.cfg')
        except FileNotFoundError:
            pass
        except  ParsingError:
            print('Error in config file:')
            print(sys.exc_info()[1])
        # project name    
        if self.config.has_option("project","name"):
            self.projectName = self.config.get("project","name")    
        elif len(sys.argv) > 1: 
            self.projectName = sys.argv[1]
        else:
            dirlist = glob.glob("*.pro")
            if len(dirlist) == 1 and os.path.isfile(dirlist[0]):
                self.projectName = dirlist[0].replace(".pro","")
            else:
                print("Please specify the project")
                exit(1)
        self.header = self.config.get("project","header",fallback = self.projectName)
        # process columns
        def proccol(s,i):
            a = s.split(':')
            res = {'name':a[0]}
            if len(a) == 1:
                res['source'] = ''
            if len(a) == 2:
                if a[1].replace('.','',1).isdigit():
                    res['width'] = int(a[1])
                    res['source'] = ''
                else:
                    res['source'] = a[1]
            if len(a) == 3:
                res['source'] = a[1]
                res['width'] = int(a[2])
            if not 'width' in res:
                if len(default_col_size) > i:
                    res['width'] = default_col_size[i]
                else:
                    res['width'] = default_col_size[-1]
            if res['source'] == '' and res['name'] != '':
                s = res['name'].lower()
                if s == 'x' or s == 'y':
                    res['source'] = s
            return res
        # columns    
        if self.config.has_section("columns"):
            default_col_size = [15,30,30,15,30]
            self.columns = self.getList("columns",proccol)
        else:
            self.columns = [{'name':"N",'source':'n','width':10},
                          {'name':"Ref",'source':'reference','width':50},
                          {'name':"Size/Package",'source':'package','width':20},
                          {'name':"Qty",'source':'quantity','width':10},
                          {'name':"Type/Value",'source':'value','width':20}]
        
        if self.config.has_section("pos_columns"):
            default_col_size = [10,15,20,20,10,10,10]
            self.pos_columns = self.getList("pos_columns",proccol)
        else:
        #Ref     Val       Package                PosX       PosY       Rot  Side
            self.pos_columns = [{'name':"N",'source':'n','width':10},
                              {'name':"Ref",'source':'reference','width':15},
                              {'name':"Val",'source':'value','width':20},
                              {'name':"Package",'source':'package','width':20},
                              {'name':"PosX",'source':'x','width':10},
                              {'name':"PosY",'source':'y','width':10},
                              {'name':"Rot",'source':'angle','width':10},
                              {'name':"Side",'source':'side','width':15}]
                              
        attr_template = r'([A-z]+)\s*\((.*)\)'

        # ignore    
        self.ignore = []    
        if self.config.has_section("ignore"):
            for i in self.config.options("ignore"):
                attr = re.match(attr_template,i.strip())
                if attr == None:
                    print("Error in ignore definition:",i)
                else:
                    self.ignore.append({'attr':attr.group(1),'match':attr.group(2)})

        # sections
        self.sections = OrderedDict()
        if self.config.has_section("sections"):
            for i in self.config.options("sections"):
                self.sections[i] = self.config.get("sections",i)
                
        # package substitutions
        self.package_sub = []        
        if self.config.has_section("packages"):
            self.package_sub = []
            for i in self.config.options("packages"):
                self.package_sub.append({'match':"^"+i+"$",'repl':self.config.get("packages",i)})
        
        # category definitions        
        self.categories = []
        if self.config.has_section("categories"):
            for i in self.config.options("categories"):
                attr = re.match(attr_template,i.strip())
                if attr == None:
                    print("Error in category definition:",i)
                else:
                    self.categories.append({'attr':attr.group(1),'match':"^"+attr.group(2)+"$",'category':self.config.get("categories",i)})
        # formats
        self.formats = {}
        self.formats["header"] = {'bold': True, 'font_size':16, 'font_color': 'navy','underline':1}
        self.formats["subheader"] = {'bold': True, 'font_size':14,'align':'left'}
        self.formats["column_header"] = {'bold': True, 'font_size':12,'align':'center','valign':'vcenter','border':2}
        self.formats["pos_number"] = {'font_size':11,'align':'center','bg_color':'#BBFF33','border':2}
        self.formats["section_header"] = {'bold': True,'font_size':12,'align':'center','border':1,'underline':1,'italic':True}
        self.formats["cell"] = {'font_size':11,'align':'center','bg_color':'#FFFFB3','border':1}
        self.formats["ref"] = {'font_size':11,'align':'left','bg_color':'#FFFFB3','border':1}
        for f in self.formats:
            if self.config.has_option("formats",f):
                try:
                    self.formats[f] = eval(self.config.get("formats",f))
                except:
                    print("Error in format ",f)
        
    def getList(self,section,proc = False):
        res = []
        n = 0
        for i in sortRef(self.config.options(section)):
            v = self.config.get(section,i)
            if proc:
                v = proc(v,n)
                n = n+1
            res.append(v)
        return res  
    
    def tryCategory(self, item):
        for i in self.categories:
            a = item.getAttr(i[0])
            if a and re.match(i[1],str(a)):
                return i[2]
        return False

           
if __name__ == '__main__':
    options = Options()
    print("Project ",options.projectName)
    brd = Board(options.projectName,options)
    brd.createXLSX()
    brd.addBOM()
    if (options.config.has_option("project","positions") and
            options.config.get("project","positions") == "yes"):
        brd.addPlacement()
    brd.writeXLSX()
