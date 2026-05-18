/**
 * 浏览器端 PPTX 文本解析（Open XML / JSZip）
 */
(function (global) {
  var A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  function extractSlideParagraphs(xmlStr) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlStr, 'application/xml');
    var paras = [];
    var current = [];
    var nodes = doc.getElementsByTagName('*');
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.localName === 't' && el.namespaceURI === A_NS) {
        current.push(el.textContent || '');
      } else if (el.localName === 'p' && el.namespaceURI === A_NS) {
        if (current.length) {
          var s = current.join('').trim();
          if (s) paras.push(s);
          current = [];
        }
      }
    }
    if (current.length) {
      var tail = current.join('').trim();
      if (tail) paras.push(tail);
    }
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
        var paras = extractSlideParagraphs(xml);
        var title = paras[0] || ('幻灯片 ' + (idx + 1));
        slides[idx] = {
          index: idx + 1,
          title: title.slice(0, 120),
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
})(typeof window !== 'undefined' ? window : this);
