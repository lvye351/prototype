/**
 * 浏览器端 PPTX 解析：按画布顺序提取文本块 + 表格（Open XML / JSZip）
 */
(function (global) {
  var A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';

  function cellText(tc) {
    var parts = [];
    var nodes = tc.getElementsByTagName('*');
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.localName === 't' && el.namespaceURI === A_NS) {
        parts.push(el.textContent || '');
      }
    }
    return parts.join('').trim();
  }

  function parseTableElement(tbl) {
    var rows = [];
    var trList = tbl.getElementsByTagName('*');
    for (var i = 0; i < trList.length; i++) {
      if (trList[i].localName !== 'tr' || trList[i].namespaceURI !== A_NS) continue;
      var tr = trList[i];
      var row = [];
      var tcList = tr.getElementsByTagName('*');
      for (var j = 0; j < tcList.length; j++) {
        if (tcList[j].localName === 'tc' && tcList[j].namespaceURI === A_NS) {
          row.push(cellText(tcList[j]));
        }
      }
      if (row.length) rows.push(row);
    }
    return rows;
  }

  function shapeParagraphs(sp) {
    var paras = [];
    var pList = sp.getElementsByTagName('*');
    for (var i = 0; i < pList.length; i++) {
      if (pList[i].localName !== 'p' || pList[i].namespaceURI !== A_NS) continue;
      var ap = pList[i];
      var parts = [];
      var tList = ap.getElementsByTagName('*');
      for (var j = 0; j < tList.length; j++) {
        if (tList[j].localName === 't' && tList[j].namespaceURI === A_NS) {
          parts.push(tList[j].textContent || '');
        }
      }
      var s = parts.join('').trim();
      if (s) paras.push(s);
    }
    return paras;
  }

  function collectFromContainer(container, blocks) {
    var children = container.children || container.childNodes;
    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      if (child.nodeType !== 1) continue;
      var tag = child.localName;
      if (tag === 'sp') {
        var paras = shapeParagraphs(child);
        if (paras.length) blocks.push({ type: 'text', paragraphs: paras });
      } else if (tag === 'graphicFrame') {
        var tbls = child.getElementsByTagName('*');
        for (var t = 0; t < tbls.length; t++) {
          if (tbls[t].localName === 'tbl' && tbls[t].namespaceURI === A_NS) {
            var rows = parseTableElement(tbls[t]);
            if (rows.length) blocks.push({ type: 'table', rows: rows });
            break;
          }
        }
      } else if (tag === 'grpSp') {
        collectFromContainer(child, blocks);
      }
    }
  }

  function extractSlideBlocks(xmlStr) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlStr, 'application/xml');
    var blocks = [];
    var trees = doc.getElementsByTagName('*');
    for (var i = 0; i < trees.length; i++) {
      if (trees[i].localName === 'spTree' && trees[i].namespaceURI === P_NS) {
        collectFromContainer(trees[i], blocks);
        break;
      }
    }
    return blocks;
  }

  function flattenParagraphs(blocks) {
    var paras = [];
    blocks.forEach(function (b) {
      if (b.type === 'text') paras = paras.concat(b.paragraphs || []);
    });
    return paras;
  }

  function slideSortKey(name) {
    var m = /slide(\d+)\.xml$/i.exec(name);
    return m ? parseInt(m[1], 10) : 0;
  }

  function buildDeckFromZip(zip, fileName) {
    var slideNames = [];
    zip.forEach(function (_path, file) {
      if (/^ppt\/slides\/slide\d+\.xml$/i.test(file.name)) slideNames.push(file.name);
    });
    slideNames.sort(function (a, b) { return slideSortKey(a) - slideSortKey(b); });

    var slides = [];
    var titleGuess = (fileName || '').replace(/\.pptx?$/i, '');

    return Promise.all(slideNames.map(function (name, idx) {
      return zip.file(name).async('string').then(function (xml) {
        var blocks = extractSlideBlocks(xml);
        var paras = flattenParagraphs(blocks);
        var title = paras[0] || ('幻灯片 ' + (idx + 1));
        slides[idx] = {
          index: idx + 1,
          title: title.slice(0, 120),
          blocks: blocks,
          paragraphs: paras,
          body: paras.slice(1, 8).join('\n')
        };
      });
    })).then(function () {
      var longTitle = '';
      if (slides[0] && slides[0].paragraphs) {
        longTitle = slides[0].paragraphs.find(function (p) {
          return p.length > 18 && /项目|汇报|方案|采购/.test(p);
        }) || slides[0].paragraphs[slides[0].paragraphs.length - 1] || '';
      }
      return {
        version: 2,
        source: fileName || 'upload.pptx',
        title: longTitle || titleGuess,
        slideCount: slides.length,
        slides: slides.filter(Boolean)
      };
    });
  }

  function parsePptxArrayBuffer(buffer, fileName) {
    if (typeof JSZip === 'undefined') {
      return Promise.reject(new Error('JSZip 未加载'));
    }
    return JSZip.loadAsync(buffer).then(function (zip) {
      return buildDeckFromZip(zip, fileName);
    });
  }

  function arrayBufferFromBase64(b64) {
    var bin = atob(b64);
    var len = bin.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }

  global.parsePptxArrayBuffer = parsePptxArrayBuffer;
  global.pptxArrayBufferFromBase64 = arrayBufferFromBase64;
  global.extractSlideBlocks = extractSlideBlocks;
})(typeof window !== 'undefined' ? window : this);
