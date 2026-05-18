/**
 * PPT 编辑页：按 blocks（文本 + 表格）渲染与回写
 */
(function (global) {
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function slideKind(slide) {
    if (slide.index === 1) return 'cover';
    if (slide.index === 3 || /目录/.test(slide.title || '')) return 'toc';
    return 'content';
  }

  function ensureBlocks(slide) {
    if (slide.blocks && slide.blocks.length) return slide.blocks;
    var paras = slide.paragraphs && slide.paragraphs.length ? slide.paragraphs : [slide.title || ''];
    return [{ type: 'text', paragraphs: paras }];
  }

  function isSectionHead(p) {
    return /^[一二三四五六七八九十百千]+、/.test(p) ||
      p === '目录' || /^方案|^第[一二三四五六七八九十]/.test(p);
  }

  function renderTableHtml(rows, bi) {
    if (!rows || !rows.length) return '';
    var maxCols = 0;
    rows.forEach(function (r) { if (r.length > maxCols) maxCols = r.length; });
    var html = '<div class="slide-table-wrap"><table class="slide-table" data-bi="' + bi + '"><tbody>';
    rows.forEach(function (row, ri) {
      html += '<tr>';
      for (var ci = 0; ci < maxCols; ci++) {
        var cell = row[ci] != null ? row[ci] : '';
        var tag = ri === 0 ? 'th' : 'td';
        html += '<' + tag + ' class="slide-cell" contenteditable="true" data-bi="' + bi + '" data-r="' + ri + '" data-c="' + ci + '">' +
          escapeHtml(cell) + '</' + tag + '>';
      }
      html += '</tr>';
    });
    return html + '</tbody></table></div>';
  }

  function renderCoverFromBlocks(slide) {
    var all = [];
    ensureBlocks(slide).forEach(function (b) {
      if (b.type === 'text') all = all.concat(b.paragraphs || []);
    });
    var sub = all[0] || '';
    var main = all.find(function (p) { return p.length > 20; }) || all[all.length - 1] || sub;
    var meta = all.find(function (p) { return /\d{4}年/.test(p); }) || all[1] || '';
    return '<div class="slide-cover">' +
      '<div class="sub slide-block" contenteditable="true">' + escapeHtml(sub) + '</div>' +
      '<div class="main-title slide-block" contenteditable="true">' + escapeHtml(main) + '</div>' +
      '<div class="meta slide-block" contenteditable="true">' + escapeHtml(meta) + '</div>' +
      '</div>';
  }

  function renderTocFromBlocks(slide) {
    var items = [];
    ensureBlocks(slide).forEach(function (b) {
      if (b.type !== 'text') return;
      (b.paragraphs || []).forEach(function (p) {
        if (p !== '目录' && !/^\d+$/.test(p)) items.push(p);
      });
    });
    var list = items.map(function (p) {
      return '<div class="slide-block" contenteditable="true">' + escapeHtml(p) + '</div>';
    }).join('');
    return '<div class="slide-toc"><h2 class="slide-block" contenteditable="true">目录</h2><div class="toc-list">' + list + '</div></div>';
  }

  function renderContentFromBlocks(slide) {
    var blocks = ensureBlocks(slide);
    var head = '';
    var bodyHtml = '';
    var headUsed = false;

    blocks.forEach(function (block, bi) {
      if (block.type === 'text') {
        (block.paragraphs || []).forEach(function (p, pi) {
          if (!headUsed && isSectionHead(p)) {
            head = p;
            headUsed = true;
            return;
          }
          bodyHtml += '<div class="slide-block" contenteditable="true" data-bi="' + bi + '" data-p="' + pi + '">' + escapeHtml(p) + '</div>';
        });
      } else if (block.type === 'table') {
        bodyHtml += '<div class="ppt-block" data-bi="' + bi + '">' + renderTableHtml(block.rows, bi) + '</div>';
      }
    });

    if (!head) {
      var ft = blocks.find(function (b) { return b.type === 'text' && b.paragraphs && b.paragraphs.length; });
      if (ft) head = ft.paragraphs[0];
    }
    if (!head) head = slide.title || '';

    return '<div class="slide-content">' +
      '<div class="head slide-block" contenteditable="true" data-role="head">' + escapeHtml(head) + '</div>' +
      '<div class="slide-blocks-body">' + bodyHtml + '</div>' +
      '</div>';
  }

  function renderSlideInner(slide) {
    var blocks = ensureBlocks(slide);
    var hasTable = blocks.some(function (b) { return b.type === 'table'; });
    var kind = slideKind(slide);

    if (hasTable || (blocks.length > 1 && kind === 'content')) return renderContentFromBlocks(slide);
    if (kind === 'cover') return renderCoverFromBlocks(slide);
    if (kind === 'toc') return renderTocFromBlocks(slide);
    return renderContentFromBlocks(slide);
  }

  function renderThumbMini(slide) {
    return '<div class="thumb-mini">' + renderSlideInner(slide) + '</div>';
  }

  function collectBlocksFromDom(stage) {
    var newBlocks = [];
    var headEl = stage.querySelector('[data-role="head"]');
    var headText = headEl ? (headEl.innerText || '').trim() : '';

    stage.querySelectorAll('.ppt-block').forEach(function (wrap) {
      var table = wrap.querySelector('table.slide-table');
      if (table) {
        var rows = [];
        table.querySelectorAll('tr').forEach(function (tr) {
          var row = [];
          tr.querySelectorAll('th, td').forEach(function (cell) {
            row.push((cell.innerText || '').trim());
          });
          if (row.length) rows.push(row);
        });
        if (rows.length) newBlocks.push({ type: 'table', rows: rows });
        return;
      }
      var paras = [];
      wrap.querySelectorAll('.slide-block').forEach(function (el) {
        if (el.getAttribute('data-role') === 'head') return;
        var t = (el.innerText || '').trim();
        if (t) paras.push(t);
      });
      if (paras.length) newBlocks.push({ type: 'text', paragraphs: paras });
    });

    if (!stage.querySelector('.ppt-block')) {
      if (stage.querySelector('.slide-cover') || stage.querySelector('.slide-toc')) {
        var parasCover = [];
        stage.querySelectorAll('.slide-block').forEach(function (el) {
          var t = (el.innerText || '').trim();
          if (t) parasCover.push(t);
        });
        if (parasCover.length) newBlocks.push({ type: 'text', paragraphs: parasCover });
      } else {
        var texts = [];
        stage.querySelectorAll('.slide-blocks-body .slide-block').forEach(function (el) {
          var t = (el.innerText || '').trim();
          if (t) texts.push(t);
        });
        if (headText || texts.length) {
          newBlocks.push({ type: 'text', paragraphs: headText ? [headText].concat(texts) : texts });
        }
        stage.querySelectorAll('.slide-table').forEach(function (table) {
          var rows = [];
          table.querySelectorAll('tr').forEach(function (tr) {
            var row = [];
            tr.querySelectorAll('th, td').forEach(function (cell) { row.push((cell.innerText || '').trim()); });
            if (row.length) rows.push(row);
          });
          if (rows.length) newBlocks.push({ type: 'table', rows: rows });
        });
      }
    } else if (headText && newBlocks.length && newBlocks[0].type === 'text' && newBlocks[0].paragraphs[0] !== headText) {
      newBlocks[0].paragraphs.unshift(headText);
    }

    return newBlocks;
  }

  function flattenParagraphs(blocks) {
    var paras = [];
    blocks.forEach(function (b) {
      if (b.type === 'text') paras = paras.concat(b.paragraphs || []);
    });
    return paras;
  }

  function syncSlideFromDom(slide, stage) {
    var blocks = collectBlocksFromDom(stage);
    if (!blocks.length) return;
    slide.blocks = blocks;
    slide.paragraphs = flattenParagraphs(blocks);
    slide.title = slide.paragraphs[0] || slide.title;
    slide.body = slide.paragraphs.slice(1, 8).join('\n');
  }

  function countTables(deck) {
    if (!deck || !deck.slides) return 0;
    var n = 0;
    deck.slides.forEach(function (s) {
      (s.blocks || []).forEach(function (b) { if (b.type === 'table') n++; });
    });
    return n;
  }

  function deckHasBlocks(deck) {
    return deck && deck.version >= 2 && deck.slides && deck.slides.length &&
      deck.slides.some(function (s) { return s.blocks && s.blocks.length; });
  }

  global.PptEditRender = {
    renderSlideInner: renderSlideInner,
    renderThumbMini: renderThumbMini,
    syncSlideFromDom: syncSlideFromDom,
    countTables: countTables,
    deckHasBlocks: deckHasBlocks
  };
})(typeof window !== 'undefined' ? window : this);
