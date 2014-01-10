/*
  MS Excel 2007 Creater v0.0.1
  Author : chuanyi.zheng@gmail.com
  History: 2012/11/07 first created
*/


(function() {
  var ContentTypes, DocPropsApp, SharedStrings, Sheet, Style, Workbook, XlRels, XlWorkbook, exec, existsSync, fs, opt, path, tool, xml,
    __bind = function(fn, me){ return function(){ return fn.apply(me, arguments); }; };

  fs = require('fs');

  path = require('path');

  exec = require('child_process');

  xml = require('xmlbuilder');

  existsSync = fs.existsSync || path.existsSync;

  tool = {
    i2a: function(i) {
      return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(i - 1);
    },
    copy: function(origin, target) {
      var f, files, oCur, s, tCur, _i, _len, _results;
      if (existsSync(origin)) {
        if (!existsSync(target)) {
          fs.mkdirSync(target, 755);
        }
        files = fs.readdirSync(origin);
        if (files) {
          _results = [];
          for (_i = 0, _len = files.length; _i < _len; _i++) {
            f = files[_i];
            oCur = origin + '/' + f;
            tCur = target + '/' + f;
            s = fs.statSync(oCur);
            if (s.isFile()) {
              _results.push(fs.writeFileSync(tCur, fs.readFileSync(oCur, ''), ''));
            } else {
              if (s.isDirectory()) {
                _results.push(tool.copy(oCur, tCur));
              } else {
                _results.push(void 0);
              }
            }
          }
          return _results;
        }
      }
    }
  };

  opt = {
    tmpl_path: __dirname
  };

  ContentTypes = (function() {
    function ContentTypes(book) {
      this.book = book;
    }

    ContentTypes.prototype.toxml = function() {
      var i, types, _i, _ref;
      types = xml.create('Types', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      types.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
      types.ele('Override', {
        PartName: '/xl/theme/theme1.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml'
      });
      types.ele('Override', {
        PartName: '/xl/styles.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
      });
      types.ele('Default', {
        Extension: 'rels',
        ContentType: 'application/vnd.openxmlformats-package.relationships+xml'
      });
      types.ele('Default', {
        Extension: 'xml',
        ContentType: 'application/xml'
      });
      types.ele('Override', {
        PartName: '/xl/workbook.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
      });
      types.ele('Override', {
        PartName: '/docProps/app.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
      });
      for (i = _i = 1, _ref = this.book.sheets.length; 1 <= _ref ? _i <= _ref : _i >= _ref; i = 1 <= _ref ? ++_i : --_i) {
        types.ele('Override', {
          PartName: '/xl/worksheets/sheet' + i + '.xml',
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
        });
      }
      types.ele('Override', {
        PartName: '/xl/sharedStrings.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
      });
      types.ele('Override', {
        PartName: '/docProps/core.xml',
        ContentType: 'application/vnd.openxmlformats-package.core-properties+xml'
      });
      return types.end();
    };

    return ContentTypes;

  })();

  DocPropsApp = (function() {
    function DocPropsApp(book) {
      this.book = book;
    }

    DocPropsApp.prototype.toxml = function() {
      var i, props, tmp, _i, _ref;
      props = xml.create('Properties', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      props.att('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
      props.att('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');
      props.ele('Application', 'Microsoft Excel');
      props.ele('DocSecurity', '0');
      props.ele('ScaleCrop', 'false');
      tmp = props.ele('HeadingPairs').ele('vt:vector', {
        size: 2,
        baseType: 'variant'
      });
      tmp.ele('vt:variant').ele('vt:lpstr', '工作表');
      tmp.ele('vt:variant').ele('vt:i4', '' + this.book.sheets.length);
      tmp = props.ele('TitlesOfParts').ele('vt:vector', {
        size: this.book.sheets.length,
        baseType: 'lpstr'
      });
      for (i = _i = 1, _ref = this.book.sheets.length; 1 <= _ref ? _i <= _ref : _i >= _ref; i = 1 <= _ref ? ++_i : --_i) {
        tmp.ele('vt:lpstr', this.book.sheets[i - 1].name);
      }
      props.ele('Company');
      props.ele('LinksUpToDate', 'false');
      props.ele('SharedDoc', 'false');
      props.ele('HyperlinksChanged', 'false');
      props.ele('AppVersion', '12.0000');
      return props.end();
    };

    return DocPropsApp;

  })();

  XlWorkbook = (function() {
    function XlWorkbook(book) {
      this.book = book;
    }

    XlWorkbook.prototype.toxml = function() {
      var i, tmp, wb, _i, _ref;
      wb = xml.create('workbook', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      wb.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
      wb.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
      wb.ele('fileVersion ', {
        appName: 'xl',
        lastEdited: '4',
        lowestEdited: '4',
        rupBuild: '4505'
      });
      wb.ele('workbookPr', {
        filterPrivacy: '1',
        defaultThemeVersion: '124226'
      });
      wb.ele('bookViews').ele('workbookView ', {
        xWindow: '0',
        yWindow: '90',
        windowWidth: '19200',
        windowHeight: '11640'
      });
      tmp = wb.ele('sheets');
      for (i = _i = 1, _ref = this.book.sheets.length; 1 <= _ref ? _i <= _ref : _i >= _ref; i = 1 <= _ref ? ++_i : --_i) {
        tmp.ele('sheet', {
          name: this.book.sheets[i - 1].name,
          sheetId: '' + i,
          'r:id': 'rId' + i
        });
      }
      wb.ele('calcPr', {
        calcId: '124519'
      });
      return wb.end();
    };

    return XlWorkbook;

  })();

  XlRels = (function() {
    function XlRels(book) {
      this.book = book;
    }

    XlRels.prototype.toxml = function() {
      var i, rs, _i, _ref;
      rs = xml.create('Relationships', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
      for (i = _i = 1, _ref = this.book.sheets.length; 1 <= _ref ? _i <= _ref : _i >= _ref; i = 1 <= _ref ? ++_i : --_i) {
        rs.ele('Relationship', {
          Id: 'rId' + i,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          Target: 'worksheets/sheet' + i + '.xml'
        });
      }
      rs.ele('Relationship', {
        Id: 'rId' + (this.book.sheets.length + 1),
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
        Target: 'theme/theme1.xml'
      });
      rs.ele('Relationship', {
        Id: 'rId' + (this.book.sheets.length + 2),
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
        Target: 'styles.xml'
      });
      rs.ele('Relationship', {
        Id: 'rId' + (this.book.sheets.length + 3),
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        Target: 'sharedStrings.xml'
      });
      return rs.end();
    };

    return XlRels;

  })();

  SharedStrings = (function() {
    function SharedStrings() {
      this.cache = {};
      this.arr = [];
    }

    SharedStrings.prototype.str2id = function(s) {
      var id;
      id = this.cache[s];
      if (id) {
        return id;
      } else {
        this.arr.push(s);
        this.cache[s] = this.arr.length;
        return this.arr.length;
      }
    };

    SharedStrings.prototype.toxml = function() {
      var i, si, sst, _i, _ref;
      sst = xml.create('sst', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      sst.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
      sst.att('count', '' + this.arr.length);
      sst.att('uniqueCount', '' + this.arr.length);
      for (i = _i = 0, _ref = this.arr.length; 0 <= _ref ? _i < _ref : _i > _ref; i = 0 <= _ref ? ++_i : --_i) {
        si = sst.ele('si');
        si.ele('t', this.arr[i]);
        si.ele('phoneticPr', {
          fontId: 1,
          type: 'noConversion'
        });
      }
      return sst.end();
    };

    return SharedStrings;

  })();

  Sheet = (function() {
    function Sheet(book, name, cols, rows) {
      var i, j, _i, _j, _ref, _ref1;
      this.book = book;
      this.name = name;
      this.cols = cols;
      this.rows = rows;
      this.data = {};
      for (i = _i = 1, _ref = this.rows; 1 <= _ref ? _i <= _ref : _i >= _ref; i = 1 <= _ref ? ++_i : --_i) {
        this.data[i] = {};
        for (j = _j = 1, _ref1 = this.cols; 1 <= _ref1 ? _j <= _ref1 : _j >= _ref1; j = 1 <= _ref1 ? ++_j : --_j) {
          this.data[i][j] = {
            v: 0,
            dataType: 'string'
          };
        }
      }
      this.merges = [];
      this.col_wd = [];
      this.row_ht = {};
      this.styles = {};
    }

    Sheet.prototype.set = function(col, row, value) {
      if ((typeof value) === "string") {
        console.log("Found string " + value);
        if ((value != null) && value !== '') {
          this.data[row][col].v = this.book.ss.str2id('' + value);
        }
        console.log("id=" + this.data[row][col].v);
        return this.data[row][col].dataType = 'string';
      } else {
        this.data[row][col].v = value;
        return this.data[row][col].dataType = 'number';
      }
    };

    Sheet.prototype.merge = function(from_cell, to_cell) {
      return this.merges.push({
        from: from_cell,
        to: to_cell
      });
    };

    Sheet.prototype.width = function(col, wd) {
      return this.col_wd.push({
        c: col,
        cw: wd
      });
    };

    Sheet.prototype.height = function(row, ht) {
      return this.row_ht[row] = ht;
    };

    Sheet.prototype.font = function(col, row, font_s) {
      return this.styles['font_' + col + '_' + row] = this.book.st.font2id(font_s);
    };

    Sheet.prototype.fill = function(col, row, fill_s) {
      return this.styles['fill_' + col + '_' + row] = this.book.st.fill2id(fill_s);
    };

    Sheet.prototype.border = function(col, row, bder_s) {
      return this.styles['bder_' + col + '_' + row] = this.book.st.bder2id(bder_s);
    };

    Sheet.prototype.align = function(col, row, align_s) {
      return this.styles['algn_' + col + '_' + row] = align_s;
    };

    Sheet.prototype.valign = function(col, row, valign_s) {
      return this.styles['valgn_' + col + '_' + row] = valign_s;
    };

    Sheet.prototype.rotate = function(col, row, textRotation) {
      return this.styles['rotate_' + col + '_' + row] = textRotation;
    };

    Sheet.prototype.wrap = function(col, row, wrap_s) {
      return this.styles['wrap_' + col + '_' + row] = wrap_s;
    };

    Sheet.prototype.style_id = function(col, row) {
      var id, inx, style;
      inx = '_' + col + '_' + row;
      style = {
        font_id: this.styles['font' + inx],
        fill_id: this.styles['fill' + inx],
        bder_id: this.styles['bder' + inx],
        align: this.styles['algn' + inx],
        valign: this.styles['valgn' + inx],
        rotate: this.styles['rotate' + inx],
        wrap: this.styles['wrap' + inx]
      };
      id = this.book.st.style2id(style);
      return id;
    };

    Sheet.prototype.toxml = function() {
      var c, cols, cw, ht, i, ix, j, m, mc, r, sd, sid, ws, _i, _j, _k, _l, _len, _len1, _ref, _ref1, _ref2, _ref3;
      ws = xml.create('worksheet', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      ws.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
      ws.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
      ws.ele('dimension', {
        ref: 'A1'
      });
      ws.ele('sheetViews').ele('sheetView', {
        workbookViewId: '0'
      });
      ws.ele('sheetFormatPr', {
        defaultRowHeight: '13.5'
      });
      if (this.col_wd.length > 0) {
        cols = ws.ele('cols');
        _ref = this.col_wd;
        for (_i = 0, _len = _ref.length; _i < _len; _i++) {
          cw = _ref[_i];
          cols.ele('col', {
            min: '' + cw.c,
            max: '' + cw.c,
            width: cw.cw,
            customWidth: '1'
          });
        }
      }
      sd = ws.ele('sheetData');
      for (i = _j = 1, _ref1 = this.rows; 1 <= _ref1 ? _j <= _ref1 : _j >= _ref1; i = 1 <= _ref1 ? ++_j : --_j) {
        r = sd.ele('row', {
          r: '' + i,
          spans: '1:' + this.cols
        });
        ht = this.row_ht[i];
        if (ht) {
          r.att('ht', ht);
          r.att('customHeight', '1');
        }
        for (j = _k = 1, _ref2 = this.cols; 1 <= _ref2 ? _k <= _ref2 : _k >= _ref2; j = 1 <= _ref2 ? ++_k : --_k) {
          ix = this.data[i][j];
          sid = this.style_id(j, i);
          if ((ix.v !== 0) || (sid !== 1)) {
            c = r.ele('c', {
              r: '' + tool.i2a(j) + i
            });
            if (sid !== 1) {
              c.att('s', '' + (sid - 1));
            }
            if (ix.v !== 0) {
              console.log(this.data[i][j]);
              if (ix.dataType === "string") {
                c.att('t', 's');
                c.ele('v', '' + (ix.v - 1));
              } else if (ix.dataType === "number") {
                c.ele('v', '' + ix.v);
              } else {
                c.ele('v', '' + ix.v);
              }
            }
          }
        }
      }
      if (this.merges.length > 0) {
        mc = ws.ele('mergeCells', {
          count: this.merges.length
        });
        _ref3 = this.merges;
        for (_l = 0, _len1 = _ref3.length; _l < _len1; _l++) {
          m = _ref3[_l];
          mc.ele('mergeCell', {
            ref: '' + tool.i2a(m.from.col) + m.from.row + ':' + tool.i2a(m.to.col) + m.to.row
          });
        }
      }
      ws.ele('phoneticPr', {
        fontId: '1',
        type: 'noConversion'
      });
      ws.ele('pageMargins', {
        left: '0.7',
        right: '0.7',
        top: '0.75',
        bottom: '0.75',
        header: '0.3',
        footer: '0.3'
      });
      ws.ele('pageSetup', {
        paperSize: '9',
        orientation: 'portrait',
        horizontalDpi: '200',
        verticalDpi: '200'
      });
      return ws.end();
    };

    return Sheet;

  })();

  Style = (function() {
    function Style(book) {
      this.book = book;
      this.cache = {};
      this.mfonts = [];
      this.mfills = [];
      this.mbders = [];
      this.mstyle = [];
      this.with_default();
    }

    Style.prototype.with_default = function() {
      this.def_font_id = this.font2id(null);
      this.def_fill_id = this.fill2id(null);
      this.def_bder_id = this.bder2id(null);
      this.def_align = '-';
      this.def_valign = '-';
      this.def_rotate = '-';
      this.def_wrap = '-';
      return this.def_style_id = this.style2id({
        font_id: this.def_font_id,
        fill_id: this.def_fill_id,
        bder_id: this.def_bder_id,
        align: this.def_align,
        valign: this.def_valign,
        rotate: this.def_rotate
      });
    };

    Style.prototype.font2id = function(font) {
      var id, k;
      font || (font = {});
      font.bold || (font.bold = '-');
      font.iter || (font.iter = '-');
      font.sz || (font.sz = '11');
      font.color || (font.color = '-');
      font.name || (font.name = 'Arial');
      font.scheme || (font.scheme = 'minor');
      font.family || (font.family = '2');
      k = 'font_' + font.bold + font.iter + font.sz + font.color + font.name + font.scheme + font.family;
      id = this.cache[k];
      if (id) {
        return id;
      } else {
        this.mfonts.push(font);
        this.cache[k] = this.mfonts.length;
        return this.mfonts.length;
      }
    };

    Style.prototype.fill2id = function(fill) {
      var id, k;
      fill || (fill = {});
      fill.type || (fill.type = 'none');
      fill.bgColor || (fill.bgColor = '-');
      fill.fgColor || (fill.fgColor = '-');
      k = 'fill_' + fill.type + fill.bgColor + fill.fgColor;
      id = this.cache[k];
      if (id) {
        return id;
      } else {
        this.mfills.push(fill);
        this.cache[k] = this.mfills.length;
        return this.mfills.length;
      }
    };

    Style.prototype.bder2id = function(bder) {
      var id, k;
      bder || (bder = {});
      bder.left || (bder.left = '-');
      bder.right || (bder.right = '-');
      bder.top || (bder.top = '-');
      bder.bottom || (bder.bottom = '-');
      k = 'bder_' + bder.left + '_' + bder.right + '_' + bder.top + '_' + bder.bottom;
      id = this.cache[k];
      if (id) {
        return id;
      } else {
        this.mbders.push(bder);
        this.cache[k] = this.mbders.length;
        return this.mbders.length;
      }
    };

    Style.prototype.style2id = function(style) {
      var id, k;
      style.align || (style.align = this.def_align);
      style.valign || (style.valign = this.def_valign);
      style.rotate || (style.rotate = this.def_rotate);
      style.wrap || (style.wrap = this.def_wrap);
      style.font_id || (style.font_id = this.def_font_id);
      style.fill_id || (style.fill_id = this.def_fill_id);
      style.bder_id || (style.bder_id = this.def_bder_id);
      k = 's_' + style.font_id + '_' + style.fill_id + '_' + style.bder_id + '_' + style.align + '_' + style.valign + '_' + style.wrap + '_' + style.rotate;
      id = this.cache[k];
      if (id) {
        return id;
      } else {
        this.mstyle.push(style);
        this.cache[k] = this.mstyle.length;
        return this.mstyle.length;
      }
    };

    Style.prototype.toxml = function() {
      var bders, cs, e, ea, es, fills, fonts, o, ss, _i, _j, _k, _l, _len, _len1, _len2, _len3, _ref, _ref1, _ref2, _ref3;
      ss = xml.create('styleSheet', {
        version: '1.0',
        encoding: 'UTF-8',
        standalone: true
      });
      ss.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
      fonts = ss.ele('fonts', {
        count: this.mfonts.length
      });
      _ref = this.mfonts;
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        o = _ref[_i];
        e = fonts.ele('font');
        if (o.bold !== '-') {
          e.ele('b');
        }
        if (o.iter !== '-') {
          e.ele('i');
        }
        e.ele('sz', {
          val: o.sz
        });
        if (o.color !== '-') {
          e.ele('color', {
            theme: o.color
          });
        }
        e.ele('name', {
          val: o.name
        });
        e.ele('family', {
          val: o.family
        });
        e.ele('charset', {
          val: '134'
        });
        if (o.scheme !== '-') {
          e.ele('scheme', {
            val: 'minor'
          });
        }
      }
      fills = ss.ele('fills', {
        count: this.mfills.length
      });
      _ref1 = this.mfills;
      for (_j = 0, _len1 = _ref1.length; _j < _len1; _j++) {
        o = _ref1[_j];
        e = fills.ele('fill');
        es = e.ele('patternFill', {
          patternType: o.type
        });
        if (o.fgColor !== '-') {
          es.ele('fgColor', {
            theme: '8',
            tint: '0.79998168889431442'
          });
        }
        if (o.bgColor !== '-') {
          es.ele('bgColor', {
            indexed: o.bgColor
          });
        }
      }
      bders = ss.ele('borders', {
        count: this.mbders.length
      });
      _ref2 = this.mbders;
      for (_k = 0, _len2 = _ref2.length; _k < _len2; _k++) {
        o = _ref2[_k];
        e = bders.ele('border');
        if (o.left !== '-') {
          e.ele('left', {
            style: o.left
          }).ele('color', {
            auto: '1'
          });
        } else {
          e.ele('left');
        }
        if (o.right !== '-') {
          e.ele('right', {
            style: o.right
          }).ele('color', {
            auto: '1'
          });
        } else {
          e.ele('right');
        }
        if (o.top !== '-') {
          e.ele('top', {
            style: o.top
          }).ele('color', {
            auto: '1'
          });
        } else {
          e.ele('top');
        }
        if (o.bottom !== '-') {
          e.ele('bottom', {
            style: o.bottom
          }).ele('color', {
            auto: '1'
          });
        } else {
          e.ele('bottom');
        }
        e.ele('diagonal');
      }
      ss.ele('cellStyleXfs', {
        count: '1'
      }).ele('xf', {
        numFmtId: '0',
        fontId: '0',
        fillId: '0',
        borderId: '0'
      }).ele('alignment', {
        vertical: 'center'
      });
      cs = ss.ele('cellXfs', {
        count: this.mstyle.length
      });
      _ref3 = this.mstyle;
      for (_l = 0, _len3 = _ref3.length; _l < _len3; _l++) {
        o = _ref3[_l];
        e = cs.ele('xf', {
          numFmtId: '0',
          fontId: o.font_id - 1,
          fillId: o.fill_id - 1,
          borderId: o.bder_id - 1,
          xfId: '0'
        });
        if (o.font_id !== 1) {
          e.att('applyFont', '1');
        }
        if (o.fill_id !== 1) {
          e.att('applyFill', '1');
        }
        if (o.bder_id !== 1) {
          e.att('applyBorder', '1');
        }
        if (o.align !== '-' || o.valign !== '-' || o.wrap !== '-') {
          e.att('applyAlignment', '1');
          ea = e.ele('alignment', {
            textRotation: (o.rotate === '-' ? '0' : o.rotate),
            horizontal: (o.align === '-' ? 'left' : o.align),
            vertical: (o.valign === '-' ? 'top' : o.valign)
          });
          if (o.wrap !== '-') {
            ea.att('wrapText', '1');
          }
        }
      }
      ss.ele('cellStyles', {
        count: '1'
      }).ele('cellStyle', {
        name: '常规',
        xfId: '0',
        builtinId: '0'
      });
      ss.ele('dxfs', {
        count: '0'
      });
      ss.ele('tableStyles', {
        count: '0',
        defaultTableStyle: 'TableStyleMedium9',
        defaultPivotStyle: 'PivotStyleLight16'
      });
      return ss.end();
    };

    return Style;

  })();

  Workbook = (function() {
    function Workbook(fpath, fname) {
      var target;
      this.fpath = fpath;
      this.fname = fname;
      this.save = __bind(this.save, this);
      this.id = '' + parseInt(Math.random() * 9999999);
      target = path.join(path.resolve(this.fpath), this.id);
      if (existsSync(target)) {
        fs.rmdirSync(target);
      }
      tool.copy(opt.tmpl_path + '/tmpl', target);
      this.sheets = [];
      this.ss = new SharedStrings;
      this.ct = new ContentTypes(this);
      this.da = new DocPropsApp(this);
      this.wb = new XlWorkbook(this);
      this.re = new XlRels(this);
      this.st = new Style(this);
    }

    Workbook.prototype.createSheet = function(name, cols, rows) {
      var sheet;
      sheet = new Sheet(this, name, cols, rows);
      this.sheets.push(sheet);
      return sheet;
    };

    Workbook.prototype.save = function(cb) {
      var args, i, opts, target, _i, _ref;
      target = path.join(path.resolve(this.fpath), this.id);
      if (!fs.existsSync(target)) {
        fs.mkdirSync(target, function(e) {
          if (!e || (e && e.code === 'EEXIST')) {
            return console.log(path + 'created');
          } else {
            return console.log(e);
          }
        });
      }
      fs.writeFileSync(target + '\\[Content_Types].xml', this.ct.toxml(), 'utf8');
      if (!fs.existsSync(path.join(target, 'docProps'))) {
        fs.mkdirSync(path.join(target, 'docProps'), function(e) {
          if (!e || (e && e.code === 'EEXIST')) {
            return console.log(path + 'created');
          } else {
            return console.log(e);
          }
        });
      }
      fs.writeFileSync(target + '\\docProps\\app.xml', this.da.toxml(), 'utf8');
      if (!fs.existsSync(path.join(target, 'xl'))) {
        fs.mkdirSync(path.join(target, 'xl'), function(e) {
          if (!e || (e && e.code === 'EEXIST')) {
            return console.log(path + 'created');
          } else {
            return console.log(e);
          }
        });
      }
      fs.writeFileSync(target + '\\xl\\workbook.xml', this.wb.toxml(), 'utf8');
      fs.writeFileSync(target + '\\xl\\sharedStrings.xml', this.ss.toxml(), 'utf8');
      fs.mkdirSync(path.join(target, path.join('xl', '_rels')), function(e) {
        if (!e || (e && e.code === 'EEXIST')) {
          return console.log(path + 'created');
        } else {
          return console.log(e);
        }
      });
      fs.writeFileSync(target + '\\xl\\_rels\\workbook.xml.rels', this.re.toxml(), 'utf8');
      fs.mkdirSync(path.join(target, path.join('xl', 'worksheets')), function(e) {
        if (!e || (e && e.code === 'EEXIST')) {
          return console.log(path + 'created');
        } else {
          return console.log(e);
        }
      });
      for (i = _i = 0, _ref = this.sheets.length; 0 <= _ref ? _i < _ref : _i > _ref; i = 0 <= _ref ? ++_i : --_i) {
        fs.writeFileSync(target + '\\xl\\worksheets\\sheet' + (i + 1) + '.xml', this.sheets[i].toxml(), 'utf8');
      }
      fs.writeFileSync(target + '\\xl\\styles.xml', this.st.toxml(), 'utf8');
      args = ' a -tzip "' + path.join(path.resolve(this.fpath), this.fname) + '" "*"';
      opts = {
        cwd: target
      };
      return exec.exec('"' + opt.tmpl_path + '\\tool\\7za.exe"' + args, opts, function(err, stdout, stderr) {
        return exec.exec('rmdir "' + target + '" /q /s', function() {
          console.log(err);
          return cb(err);
        });
      });
    };

    Workbook.prototype.cancel = function() {
      return fs.rmdirSync(target);
    };

    return Workbook;

  })();

  module.exports = {
    createWorkbook: function(fpath, fname) {
      return new Workbook(fpath, fname);
    }
  };

}).call(this);
