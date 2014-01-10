###
  MS Excel 2007 Creater v0.0.1
  Author : chuanyi.zheng@gmail.com
  History: 2012/11/07 first created
###

fs  = require 'fs'
path = require 'path'
exec = require 'child_process'
xml = require 'xmlbuilder'
existsSync = fs.existsSync || path.existsSync

tool =
  i2a : (i) ->
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(i-1)

  copy : (origin, target) ->
  	if existsSync(origin)
      fs.mkdirSync(target, 755) if not existsSync(target)
      files = fs.readdirSync(origin)
      if files
        for f in files
          oCur = origin + '/' + f
          tCur = target + '/' + f
          s = fs.statSync(oCur)
          if s.isFile()
            fs.writeFileSync(tCur,fs.readFileSync(oCur,''),'')
          else
            if s.isDirectory()
              tool.copy oCur, tCur

opt = 
  tmpl_path : __dirname

class ContentTypes
  constructor: (@book)->

  toxml:()->
    types = xml.create('Types',{version:'1.0',encoding:'UTF-8',standalone:true})
    types.att('xmlns','http://schemas.openxmlformats.org/package/2006/content-types')
    types.ele('Override',{PartName:'/xl/theme/theme1.xml',ContentType:'application/vnd.openxmlformats-officedocument.theme+xml'})
    types.ele('Override',{PartName:'/xl/styles.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    types.ele('Default',{Extension:'rels',ContentType:'application/vnd.openxmlformats-package.relationships+xml'})
    types.ele('Default',{Extension:'xml',ContentType:'application/xml'})
    types.ele('Override',{PartName:'/xl/workbook.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    types.ele('Override',{PartName:'/docProps/app.xml',ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    for i in [1..@book.sheets.length]
      types.ele('Override',{PartName:'/xl/worksheets/sheet'+i+'.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
    types.ele('Override',{PartName:'/xl/sharedStrings.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
    types.ele('Override',{PartName:'/docProps/core.xml',ContentType:'application/vnd.openxmlformats-package.core-properties+xml'})
    return types.end()

class DocPropsApp
  constructor: (@book)->

  toxml: ()->
    props = xml.create('Properties',{version:'1.0',encoding:'UTF-8',standalone:true})
    props.att('xmlns','http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
    props.att('xmlns:vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
    props.ele('Application','Microsoft Excel')
    props.ele('DocSecurity','0')
    props.ele('ScaleCrop','false')
    tmp = props.ele('HeadingPairs').ele('vt:vector',{size:2,baseType:'variant'})
    tmp.ele('vt:variant').ele('vt:lpstr','工作表')
    tmp.ele('vt:variant').ele('vt:i4',''+@book.sheets.length)
    tmp = props.ele('TitlesOfParts').ele('vt:vector',{size:@book.sheets.length,baseType:'lpstr'})
    for i in [1..@book.sheets.length]
      tmp.ele('vt:lpstr',@book.sheets[i-1].name)
    props.ele('Company')
    props.ele('LinksUpToDate','false')
    props.ele('SharedDoc','false')  
    props.ele('HyperlinksChanged','false')  
    props.ele('AppVersion','12.0000') 
    return props.end()

class XlWorkbook
  constructor: (@book)->

  toxml: ()->
    wb = xml.create('workbook',{version:'1.0',encoding:'UTF-8',standalone:true})
    wb.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    wb.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    wb.ele('fileVersion ',{appName:'xl',lastEdited:'4',lowestEdited:'4',rupBuild:'4505'})
    wb.ele('workbookPr',{filterPrivacy:'1',defaultThemeVersion:'124226'}) 
    wb.ele('bookViews').ele('workbookView ',{xWindow:'0',yWindow:'90',windowWidth:'19200',windowHeight:'11640'})
    tmp = wb.ele('sheets')
    for i in [1..@book.sheets.length]
      tmp.ele('sheet',{name:@book.sheets[i-1].name,sheetId:''+i,'r:id':'rId'+i})
    wb.ele('calcPr',{calcId:'124519'})
    return wb.end()

class XlRels
  constructor: (@book)->
  
  toxml: ()->
    rs = xml.create('Relationships',{version:'1.0',encoding:'UTF-8',standalone:true})
    rs.att('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
    for i in [1..@book.sheets.length]
      rs.ele('Relationship',{Id:'rId'+i,Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target:'worksheets/sheet'+i+'.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+1),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',Target:'theme/theme1.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+2),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',Target:'styles.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+3),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',Target:'sharedStrings.xml'})
    return rs.end()

class SharedStrings
  constructor: ()->
    @cache = {}
    @arr = []

  str2id: (s)->
    id = @cache[s]
    if id
      return id
    else
      @arr.push s
      @cache[s] = @arr.length
      return @arr.length

  toxml: ()->
    sst = xml.create('sst',{version:'1.0',encoding:'UTF-8',standalone:true})
    sst.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    sst.att('count',''+@arr.length)
    sst.att('uniqueCount',''+@arr.length)
    for i in [0...@arr.length]
      si = sst.ele('si')
      si.ele('t',@arr[i])
      si.ele('phoneticPr',{fontId:1,type:'noConversion'})
    return sst.end()

class Sheet
  constructor: (@book, @name, @cols, @rows) ->
    @data = {}
    for i in [1..@rows]
      @data[i] = {}
      for j in [1..@cols]
        @data[i][j] = {v:0, dataType: 'string'}
    @merges = []
    @col_wd = []
    @row_ht = {}
    @styles = {}

  set: (col, row, value) ->
    if (typeof value) is "string"
      console.log "Found string " + value
      @data[row][col].v = @book.ss.str2id(''+value) if value? and value isnt ''
      console.log "id=" + @data[row][col].v
      @data[row][col].dataType = 'string'
    else
      @data[row][col].v = value
      @data[row][col].dataType = 'number'

  merge: (from_cell, to_cell) ->
    @merges.push({from:from_cell, to:to_cell})

  width: (col, wd) ->
    @col_wd.push {c:col,cw:wd}

  height: (row, ht) ->
    @row_ht[row] = ht

  font: (col, row, font_s)->
    @styles['font_'+col+'_'+row] = @book.st.font2id(font_s)

  fill: (col, row, fill_s)->
    @styles['fill_'+col+'_'+row] = @book.st.fill2id(fill_s)

  border: (col, row, bder_s)->
    @styles['bder_'+col+'_'+row] = @book.st.bder2id(bder_s)

  align: (col, row, align_s)->
    @styles['algn_'+col+'_'+row] = align_s

  valign: (col, row, valign_s)->
    @styles['valgn_'+col+'_'+row] = valign_s

  rotate: (col, row, textRotation)->
    @styles['rotate_'+col+'_'+row] = textRotation

  wrap: (col, row, wrap_s)->
    @styles['wrap_'+col+'_'+row] = wrap_s

  style_id: (col, row) ->
    inx = '_'+col+'_'+row
    style = {
      font_id:@styles['font'+inx],
      fill_id:@styles['fill'+inx],
      bder_id:@styles['bder'+inx],
      align:@styles['algn'+inx],
      valign:@styles['valgn'+inx],
      rotate:@styles['rotate'+inx],
      wrap:@styles['wrap'+inx]
    }
    id = @book.st.style2id(style)
    return id

  toxml: () ->
    ws = xml.create('worksheet',
                    {version:'1.0',encoding:'UTF-8',standalone:true})
    ws.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ws.att('xmlns:r',
           'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ws.ele('dimension',{ref:'A1'})
    ws.ele('sheetViews').ele('sheetView',{workbookViewId:'0'})
    ws.ele('sheetFormatPr',{defaultRowHeight:'13.5'})
  
    if @col_wd.length > 0
      cols = ws.ele('cols')
      for cw in @col_wd
        cols.ele('col',{min:''+cw.c,max:''+cw.c,width:cw.cw,customWidth:'1'})
    sd = ws.ele('sheetData')
    for i in [1..@rows]
      r = sd.ele('row',{r:''+i,spans:'1:'+@cols})
      ht = @row_ht[i]
      if ht
        r.att('ht',ht)
        r.att('customHeight','1')
      for j in [1..@cols]
        ix = @data[i][j]
        sid = @style_id(j, i)
        if (ix.v isnt 0) or (sid isnt 1)
          c = r.ele('c',{r:''+tool.i2a(j)+i})
          c.att('s',''+(sid-1)) if sid isnt 1
          if ix.v isnt 0
            # comment out for test generating number
            console.log @data[i][j]
            if ix.dataType is "string"
              
              c.att('t','s')
              c.ele('v', '' + (ix.v-1))
            else if ix.dataType is "number"
              c.ele('v', '' + ix.v)
            else
              c.ele('v', '' + ix.v)
    if @merges.length > 0
      mc = ws.ele('mergeCells',{count:@merges.length})
      for m in @merges
        mc.ele('mergeCell',
               {ref:(''+tool.i2a(m.from.col)+m.from.row+':'+tool.i2a(m.to.col)+m.to.row)})
    ws.ele('phoneticPr',{fontId:'1',type:'noConversion'})
    ws.ele('pageMargins',{left:'0.7',right:'0.7',top:'0.75',bottom:'0.75',header:'0.3',footer:'0.3'})
    ws.ele('pageSetup',{paperSize:'9',orientation:'portrait',horizontalDpi:'200',verticalDpi:'200'})
    return ws.end()

class Style
  constructor: (@book)->
    @cache = {}
    @mfonts = []  # font style
    @mfills = []  # fill style
    @mbders = []  # border style
    @mstyle = []  # cell style<ref-font,ref-fill,ref-border,align>
    @with_default()

  with_default:()->
    @def_font_id = @font2id(null)
    @def_fill_id = @fill2id(null)
    @def_bder_id = @bder2id(null)
    @def_align = '-'
    @def_valign = '-'
    @def_rotate = '-'
    @def_wrap = '-'
    @def_style_id = @style2id({
      font_id:@def_font_id,
      fill_id:@def_fill_id,
      bder_id:@def_bder_id,
      align:@def_align,
      valign:@def_valign,
      rotate:@def_rotate})

  font2id: (font)->
    font or= {}
    font.bold or= '-'
    font.iter or= '-'
    font.sz or= '11'
    font.color or= '-'
    font.name or= 'Arial'
    font.scheme or='minor'
    font.family or= '2'
    k = 'font_'+font.bold+font.iter+font.sz+font.color+font.name+font.scheme+font.family
    id = @cache[k]
    if id
      return id
    else
      @mfonts.push font
      @cache[k] = @mfonts.length
      return @mfonts.length

  fill2id: (fill)->
    fill or= {}
    fill.type or= 'none'
    fill.bgColor or= '-'
    fill.fgColor or= '-'
    k = 'fill_' + fill.type + fill.bgColor + fill.fgColor
    id = @cache[k]
    if id
      return id
    else
      @mfills.push fill
      @cache[k] = @mfills.length
      return @mfills.length

  bder2id: (bder)->
    bder or= {}
    bder.left or= '-'
    bder.right or= '-'
    bder.top or= '-'
    bder.bottom or= '-'
    k = 'bder_'+bder.left+'_'+bder.right+'_'+bder.top+'_'+bder.bottom
    id = @cache[k]
    if id
      return id
    else
      @mbders.push bder
      @cache[k] = @mbders.length
      return @mbders.length

  style2id:(style)->
    style.align or= @def_align
    style.valign or= @def_valign
    style.rotate or= @def_rotate
    style.wrap or= @def_wrap
    style.font_id or= @def_font_id
    style.fill_id or= @def_fill_id
    style.bder_id or= @def_bder_id
    k = 's_' + style.font_id + '_' + style.fill_id + '_' + style.bder_id + '_' + style.align + '_' + style.valign + '_' + style.wrap + '_' + style.rotate
    id = @cache[k]
    if id
      return id
    else
      @mstyle.push style
      @cache[k] = @mstyle.length
      return @mstyle.length

  toxml: ()->
    ss = xml.create('styleSheet',{version:'1.0',encoding:'UTF-8',standalone:true})
    ss.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    fonts = ss.ele('fonts',{count:@mfonts.length})
    for o in @mfonts
      e = fonts.ele('font')
      e.ele('b') if o.bold isnt '-'
      e.ele('i') if o.iter isnt '-'
      e.ele('sz',{val:o.sz})
      e.ele('color',{theme:o.color}) if o.color isnt '-'
      e.ele('name',{val:o.name})
      e.ele('family',{val:o.family})
      e.ele('charset',{val:'134'})
      e.ele('scheme',{val:'minor'}) if o.scheme isnt '-'
    fills = ss.ele('fills',{count:@mfills.length})
    for o in @mfills
      e = fills.ele('fill')
      es = e.ele('patternFill',{patternType:o.type})
      es.ele('fgColor',{theme:'8',tint:'0.79998168889431442'}) if o.fgColor isnt '-'
      es.ele('bgColor',{indexed:o.bgColor}) if o.bgColor isnt '-'
    bders = ss.ele('borders',{count:@mbders.length})
    for o in @mbders
      e = bders.ele('border')
      if o.left isnt '-' then e.ele('left',{style:o.left}).ele('color',{auto:'1'}) else e.ele('left')
      if o.right isnt '-' then e.ele('right',{style:o.right}).ele('color',{auto:'1'}) else e.ele('right')
      if o.top isnt '-' then e.ele('top',{style:o.top}).ele('color',{auto:'1'}) else e.ele('top')
      if o.bottom isnt '-' then e.ele('bottom',{style:o.bottom}).ele('color',{auto:'1'}) else e.ele('bottom')
      e.ele('diagonal')
    ss.ele('cellStyleXfs',{count:'1'}).ele('xf',{numFmtId:'0',fontId:'0',fillId:'0',borderId:'0'}).ele('alignment',{vertical:'center'})
    cs = ss.ele('cellXfs',{count:@mstyle.length})
    for o in @mstyle
      e = cs.ele('xf',{numFmtId:'0',fontId:(o.font_id-1),fillId:(o.fill_id-1),borderId:(o.bder_id-1),xfId:'0'})
      e.att('applyFont','1') if o.font_id isnt 1
      e.att('applyFill','1') if o.fill_id isnt 1
      e.att('applyBorder','1') if o.bder_id isnt 1
      if o.align isnt '-' or o.valign isnt '-' or o.wrap isnt '-'
        e.att('applyAlignment','1')
        ea = e.ele('alignment',{textRotation:(if o.rotate is '-' then '0' else o.rotate),horizontal:(if o.align is '-' then 'left' else o.align), vertical:(if o.valign is '-' then 'top' else o.valign)})
        ea.att('wrapText','1') if o.wrap isnt '-'
    ss.ele('cellStyles',{count:'1'}).ele('cellStyle',{name:'常规',xfId:'0',builtinId:'0'})
    ss.ele('dxfs',{count:'0'})
    ss.ele('tableStyles',{count:'0',defaultTableStyle:'TableStyleMedium9',defaultPivotStyle:'PivotStyleLight16'})
    return ss.end()

class Workbook
  constructor: (@fpath, @fname) ->
    @id = ''+parseInt(Math.random()*9999999)
    
    # create temp folder & copy template data
    target = path.join(path.resolve(@fpath),@id)
    fs.rmdirSync(target) if existsSync(target)
    tool.copy (opt.tmpl_path + '/tmpl'),target
    # init
    @sheets = []
    @ss = new SharedStrings
    @ct = new ContentTypes(@)
    @da = new DocPropsApp(@)
    @wb = new XlWorkbook(@)
    @re = new XlRels(@)
    @st = new Style(@)

  createSheet: (name, cols, rows) ->
    sheet = new Sheet(@,name,cols,rows)
    @sheets.push sheet
    return sheet

  save: (cb) =>
    target = path.join(path.resolve(@fpath),@id)

    # 1 - build [Content_Types].xml
    if not fs.existsSync(target)
      fs.mkdirSync target, (e) ->
        if !e or (e and e.code is 'EEXIST')
            console.log path + 'created'
        else
            console.log(e)
      
    fs.writeFileSync(target+'\\[Content_Types].xml',@ct.toxml(),'utf8')
    
    # 2 - build docProps/app.xml
    if not fs.existsSync(path.join(target,'docProps'))
      fs.mkdirSync path.join(target,'docProps'), (e) ->
        if !e or (e and e.code is 'EEXIST')
            console.log path + 'created'
        else
            console.log(e)
    fs.writeFileSync(target+'\\docProps\\app.xml',@da.toxml(),'utf8')
    
    # 3 - build xl/workbook.xml
    if not fs.existsSync(path.join(target,'xl'))
      fs.mkdirSync path.join(target,'xl'), (e) ->
        if !e or (e and e.code is 'EEXIST')
            console.log path + 'created'
        else
            console.log(e)
    fs.writeFileSync(target+'\\xl\\workbook.xml',@wb.toxml(),'utf8')
    
    # 4 - build xl/sharedStrings.xml
    fs.writeFileSync(target+'\\xl\\sharedStrings.xml',@ss.toxml(),'utf8')
    
    # 5 - build xl/_rels/workbook.xml.rels
    fs.mkdirSync path.join(target,path.join('xl','_rels')), (e) ->
        if !e or (e and e.code is 'EEXIST')
            console.log path + 'created'
        else
            console.log(e)
    fs.writeFileSync(target+'\\xl\\_rels\\workbook.xml.rels',@re.toxml(),'utf8')
    
    # 6 - build xl/worksheets/sheet(1-N).xml
    fs.mkdirSync path.join(target,path.join('xl', 'worksheets')), (e) ->
        if !e or (e and e.code is 'EEXIST')
            console.log path + 'created'
        else
            console.log(e)
    for i in [0...@sheets.length]
      fs.writeFileSync(target+'\\xl\\worksheets\\sheet'+(i+1)+'.xml',@sheets[i].toxml(),'utf8')
      
    # 7 - build xl/styles.xml
    fs.writeFileSync(target+'\\xl\\styles.xml',@st.toxml(),'utf8')    
    
    # 8 - compress temp folder to target file
    args = ' a -tzip "' + path.join(path.resolve(@fpath),@fname) + '" "*"'
    opts = {cwd:target}

    exec.exec '"'+opt.tmpl_path+'\\tool\\7za.exe"' + args, opts, (err,stdout,stderr)->
      
      # 9 - delete temp folder
      exec.exec 'rmdir "' + target + '" /q /s',()->
        console.log(err)
        cb err

  cancel: () ->
    #target = path.join(path.resolve(@fpath),@id)
    # delete temp folder
    fs.rmdirSync target

module.exports = 
  createWorkbook: (fpath, fname)->
    return new Workbook(fpath, fname)
