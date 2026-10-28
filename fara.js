/*
 * Intencje z FARY → wykaz parafii (Gosławice)
 * Rdzeń bez zależności od Office.js: czytanie eksportu FARY (docx),
 * budowa wykazu tygodnia i zamiana na OOXML ze stylami szablonu.
 * Działa w przeglądarce (dodatek) i w Node (testy).
 */
(function (root) {
  'use strict';

  var DAY = 86400000;
  var DNI = ['Niedziela', 'Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota'];

  var DEFAULTS = {
    uwagaTygodnia: 'Od poniedziałku do czwartku od 17.30 do 17.55 – Adoracja w ciszy i okazja do spowiedzi.',
    srodaWotywa: true,
    piatekAdoracja: true,
    majPazdziernik: true,
    pierwszeDniMiesiaca: true,
    kolektaPole: true
  };

  // ---------- daty ----------
  function ud(y, m, d) { return Date.UTC(y, m - 1, d); }
  function addDays(t, n) { return t + n * DAY; }
  function wd(t) { return new Date(t).getUTCDay(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmt(t) { var d = new Date(t); return pad(d.getUTCDate()) + '.' + pad(d.getUTCMonth() + 1) + '.' + d.getUTCFullYear(); }

  function roman(n) {
    var v = [[10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']], s = '';
    for (var i = 0; i < v.length; i++) while (n >= v[i][0]) { s += v[i][1]; n -= v[i][0]; }
    return s;
  }

  function easter(y) {
    var a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4,
      f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30,
      i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
    var month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
    return ud(y, month, day);
  }

  // Tydzień liturgiczny dla tygodnia pon.–niedz. = tydzień niedzieli poprzedzającej poniedziałek.
  function nazwaTygodnia(monday) {
    var S = addDays(monday, -1), y = new Date(S).getUTCFullYear();
    var E = easter(y);
    var xmas = ud(y, 12, 25), w = wd(xmas);
    var A1 = addDays(xmas, w === 0 ? -28 : -w - 21), CK = addDays(A1, -7);
    var jan6 = ud(new Date(monday).getUTCFullYear(), 1, 6);
    var B = wd(jan6) === 0 ? addDays(jan6, 1) : addDays(jan6, 7 - wd(jan6)); // Chrzest Pański
    var wk = function (a, b) { return Math.round((a - b) / DAY / 7); };

    if (S >= xmas || monday < B) return 'Okres Bożego Narodzenia';
    if (S >= A1) return roman(wk(S, A1) + 1) + ' Tydzień Adwentu';
    if (S < addDays(E, -42)) return roman(Math.floor((monday - B) / DAY / 7) + 1) + ' Tydzień Zwykły';
    if (S < addDays(E, -7)) return roman(wk(S, addDays(E, -42)) + 1) + ' Tydzień Wielkiego Postu';
    if (S === addDays(E, -7)) return 'Wielki Tydzień';
    if (S === E) return 'Oktawa Wielkiej Nocy';
    if (S < addDays(E, 49)) return roman(wk(S, E) + 1) + ' Tydzień Wielkanocny';
    return roman(34 - wk(CK, S)) + ' Tydzień Zwykły';
  }

  // ---------- czytanie FARY ----------
  function textOf(node) {
    var out = '';
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.nodeName === 'w:t') out += c.textContent;
        else if (c.nodeName === 'w:tab') out += '\t';
        else if (c.nodeName === 'w:br' || c.nodeName === 'w:cr') out += '\n';
        else walk(c);
      }
    })(node);
    return out;
  }
  function children(n, name) {
    var r = [];
    for (var c = n.firstChild; c; c = c.nextSibling) if (c.nodeName === name) r.push(c);
    return r;
  }
  function cellLines(tc) {
    var lines = [];
    children(tc, 'w:p').forEach(function (p) {
      textOf(p).split('\n').forEach(function (l) {
        l = l.replace(/\s+/g, ' ').trim();
        if (l) lines.push(l);
      });
    });
    return lines;
  }

  var RE_DATA = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;

  // Zwraca listę dni: {date, name, lines[], rows:[{hour, lines[]}]}
  function parseFaraXml(doc) {
    var days = [], cur = null;
    var rows = doc.getElementsByTagName('w:tr');
    for (var i = 0; i < rows.length; i++) {
      var tcs = children(rows[i], 'w:tc');
      if (tcs.length === 1) {
        var ls = cellLines(tcs[0]);
        var m = ls.length && ls[0].match(RE_DATA);
        if (!m) continue; // tytuł „INTENCJE MSZY ŚW. …”
        cur = {
          date: ud(+m[3], +m[2], +m[1]),
          name: ls[0].slice(0, m.index).trim(),
          lines: ls.slice(1),
          rows: []
        };
        days.push(cur);
      } else if (tcs.length >= 2 && cur) {
        var hour = cellLines(tcs[0]).join(' ');
        // każdy akapit komórki to osobna intencja/dopisek; linia od małej litery to ciąg dalszy poprzedniej
        var items = [];
        cellLines(tcs[tcs.length - 1]).forEach(function (l) {
          var prev = items[items.length - 1];
          if (prev && /^[a-ząćęłńóśźż]/.test(l)) items[items.length - 1] = prev + ' ' + l;
          else items.push(l);
        });
        cur.rows.push({ hour: hour, lines: items });
      }
    }
    if (!days.length) throw new Error('To nie wygląda na wydruk intencji z FARY (brak tabeli z dniami).');
    return days;
  }

  // ---------- reguły tekstu ----------
  var RE_DODATEK = /^(po mszy|w czasie mszy|podczas mszy|na zakończenie|przed mszą|po nabożeństwie|po różańcu)/i;
  var RE_SWIETY = /^(św\.|bł\.|świętych|najświętszej|nmp|matki bożej|niepokalanego|najśw)/i;

  function godzina(h) {
    var m = String(h).match(/^(\d{1,2})[:.](\d{2})$/);
    return m ? (+m[1]) + '.' + m[2] : String(h).trim();
  }
  function godzinaNum(h) { var m = String(h).match(/(\d{1,2})[:.](\d{2})/); return m ? +m[1] * 60 + +m[2] : -1; }

  function wewnatrzNawiasu(s, i) {
    var o = 0;
    for (var k = 0; k < i; k++) { if (s[k] === '(') o++; else if (s[k] === ')') o--; }
    return o > 0;
  }

  // Dopiski okolicznościowe w nawias: „z okazji …”, „w (kolejną|1.) rocznicę …”, „w 30. dzień po śmierci”
  var RE_NAWIAS = /z okazji(?=\s)|\bw (kolejną |kolejne |\d+\.? )?rocznicę(?=\s)|\bw (\d+\.? )?dzień (po śmierci|urodzin)/gi;
  function dopiskiWNawias(s) {
    var re = new RegExp(RE_NAWIAS.source, 'gi'), m, out = s;
    while ((m = re.exec(out))) {
      var start = m.index;
      if (wewnatrzNawiasu(out, start)) continue;
      var rest = out.slice(start), stop = rest.search(/,|;|\(|\)|(?<!\d)\.|\s(oraz|a także|i za|i o|w intencji)\s/);
      var seg = stop > 0 ? rest.slice(0, stop) : rest, phrase = seg.replace(/\s+$/, '');
      out = out.slice(0, start) + '(' + phrase + ')' + out.slice(start + phrase.length);
      re.lastIndex = start + phrase.length + 2;
    }
    return out;
  }

  function czyscIntencje(t) {
    t = t.replace(/\s+/g, ' ')
      .replace(/\(\s+/g, '(').replace(/\s+\)/g, ')')
      .replace(/\s+([,.;:])/g, '$1')
      .replace(/(\d)\.(?=[a-ząćęłńóśźż])/gi, '$1. ')
      .replace(/\bz ok\.\s*/gi, 'z okazji ')
      .replace(/(^|\s)zmarł(ego|ej|ą)(?=\s)/gi, '$1+')
      .replace(/(^|\s)zmarłych(?=\s)/gi, '$1++')
      .replace(/[\/\\]+(?=[.\s]*$)/, '')
      .trim();
    t = dopiskiWNawias(t).replace(/,\s*\(/g, ' (');
    t = t.replace(/[,;\s-]+$/, '');
    if (!/[.!?…:]$/.test(t)) t += '.';
    return t;
  }

  // Tekst intencji → fragmenty: nawiasy kursywą, intencja za parafian pogrubiona
  function fragmentyIntencji(t) {
    var bold = /w intencji parafian/i.test(t), out = [], last = 0, re = /\([^()]*\)/g, m;
    while ((m = re.exec(t))) {
      if (m.index > last) out.push({ t: t.slice(last, m.index), b: bold });
      out.push({ t: m[0], i: true, b: bold });
      last = m.index + m[0].length;
    }
    if (last < t.length) out.push({ t: t.slice(last), b: bold });
    return out;
  }

  // Wspomnienia/święta z drugiej linii nagłówka dnia
  function dopiskiDnia(day, niedziela) {
    var items = [];
    day.lines.forEach(function (l) { l.split(';').forEach(function (x) { x = x.trim(); if (x) items.push(x); }); });
    var naglowek = null, sygnaturki = [];
    items.forEach(function (it) {
      if (it.toLowerCase() === day.name.toLowerCase()) return;
      var swieto = /^(święto|uroczystość)(?=\s)/i.test(it);
      var wsp = /^wspomnienie\b/i.test(it);
      if (niedziela) {
        // niedziela ma pierwszeństwo przed wspomnieniami i świętami; uroczystość zostaje
        if (wsp || RE_SWIETY.test(it) || /^święto(?=\s)/i.test(it)) return;
        if (naglowek === null) naglowek = { t: it, sep: !/^uroczystość(?=\s)/i.test(it) };
        else sygnaturki.push(it);
        return;
      }
      if (naglowek) return; // w dzień powszedni tylko jedno: wspomnienie obowiązkowe, święto albo uroczystość
      if (swieto) naglowek = { t: it, swieto: true };
      else if (wsp) naglowek = { t: /^wspomnienie wszystkich/i.test(it) ? it : it.replace(/^wspomnienie\s+/i, '') };
      // wspomnienia dowolne i inne dopiski z FARY pomijamy
    });
    return { naglowek: naglowek, sygnaturki: sygnaturki };
  }

  // ---------- budowa tygodnia ----------
  function P(style, runs) { return { style: style, runs: runs }; }
  var ST_DZIEN = 'Dzień tygodnia', ST_SYG = 'Kolekta i sygnaturki', ST_TXT = 'Bez odstępów1';

  function msze(day) {
    var list = [], pre = [];
    day.rows.forEach(function (r) {
      var lines = r.lines.slice();
      if (r.hour) {
        var first = lines.length && !RE_DODATEK.test(lines[0]) ? lines.shift() : 'Wolna intencja';
        list.push({ hour: r.hour, intencje: [first], dodatki: [] });
      }
      lines.forEach(function (l) {
        var cur = list[list.length - 1];
        if (RE_DODATEK.test(l)) (cur ? cur.dodatki : pre).push(l);
        else if (cur) cur.intencje.push(l);
        else pre.push(l);
      });
    });
    return { list: list, pre: pre };
  }

  function dodatek(t) { return P(ST_TXT, [{ t: '\t' }, { t: t.replace(/\s+/g, ' ').trim(), i: true }]); }

  function budujDzien(day, o) {
    var out = [], w = wd(day.date), dz = new Date(day.date), dom = dz.getUTCDate(), mies = dz.getUTCMonth() + 1;
    var niedziela = w === 0;
    var dop = dopiskiDnia(day, niedziela);

    // nagłówek dnia
    var nazwa = day.name || DNI[w];
    if (niedziela && !/niedziela/i.test(nazwa)) nazwa = 'Niedziela ' + fmt(day.date) + ' ' + nazwa;
    else nazwa = nazwa + ' ' + fmt(day.date);
    if (dop.naglowek) nazwa += (dop.naglowek.sep ? ' | ' : ' ') + dop.naglowek.t;
    if (o.pierwszeDniMiesiaca && dom <= 7 && w >= 4 && w <= 6) nazwa += ' | I ' + DNI[w] + ' miesiąca';
    if (o.pierwszeDniMiesiaca && dom === 13 && mies >= 5 && mies <= 10) nazwa += ' | Dzień Fatimski';
    out.push(P(ST_DZIEN, [{ t: nazwa }]));

    if (niedziela) {
      dop.sygnaturki.forEach(function (s) { out.push(P(ST_SYG, [{ t: s }])); });
      if (o.kolektaPole) out.push(P(ST_SYG, [{ t: 'Kolekta …', hl: 'yellow' }]));
    }

    var m = msze(day);
    m.pre.forEach(function (t) { out.push(dodatek(t)); });
    if (o.srodaWotywa && w === 3 && m.list.length && !(dop.naglowek && dop.naglowek.swieto)) out.push(dodatek('Wotywna o św. Józefie'));

    var wieczorna = -1;
    if (!niedziela) m.list.forEach(function (ms, i) { if (godzinaNum(ms.hour) >= 17 * 60) wieczorna = i; });

    m.list.forEach(function (ms, i) {
      var wiele = ms.intencje.length > 1;
      ms.intencje.forEach(function (t, k) {
        var runs = k === 0 ? [{ t: godzina(ms.hour), b: true }, { t: '\t' }] : [{ t: '\t' }];
        if (wiele) runs.push({ t: (k + 1) + '. ', b: true });
        var tekst = /^wolna intencja/i.test(t) ? 'Wolna intencja' : czyscIntencje(t);
        out.push(P(ST_TXT, runs.concat(fragmentyIntencji(tekst))));
      });
      var dod = ms.dodatki.slice();

      if (o.srodaWotywa && w === 3 && i === 0 && !(dop.naglowek && dop.naglowek.swieto) && !dod.some(function (d) { return /litani/i.test(d); }))
        dod.push('Po Mszy św. litania do św. Józefa');

      if (i === wieczorna) {
        var nab = null;
        if (o.majPazdziernik && mies === 5) nab = 'nabożeństwo majowe';
        if (o.majPazdziernik && mies === 10) nab = 'różaniec';
        if (nab && dod.some(function (d) { return /różaniec|majowe/i.test(d); })) nab = null;
        var ador = o.piatekAdoracja && w === 5 && !dod.some(function (d) { return /adoracj/i.test(d); });
        if (ador && !nab && dod.length) {
          dod[dod.length - 1] = dod[dod.length - 1].replace(/[.\s]+$/, '') + ' i adoracja do godz. 19.30';
          ador = false;
        }
        if (nab || ador) {
          var t = 'Po Mszy św. ' + (nab ? nab : '') + (nab && ador ? ' i ' : '') + (ador ? 'adoracja do godz. 19.30' : '');
          dod.push(nab && !ador ? t + '.' : t);
        }
      }
      dod.forEach(function (d) { out.push(dodatek(d)); });
    });
    return out;
  }

  // Nazwa pliku jak dotąd: I_Z30_26.10.2026, I_WP3_…, I_WLK2_…
  function arabic(r) { var v = { I: 1, V: 5, X: 10 }, n = 0; for (var i = 0; i < r.length; i++) { var a = v[r[i]], b = v[r[i + 1]] || 0; n += a < b ? -a : a; } return n; }
  function nazwaPliku(nazwa, mon) {
    var m = nazwa.match(/^([IVX]+) Tydzień (\S+)/), kod = 'X';
    if (m) kod = ({ 'Zwykły': 'Z', 'Wielkiego': 'WP', 'Wielkanocny': 'WLK', 'Adwentu': 'A' })[m[2]] + arabic(m[1]);
    else if (nazwa === 'Wielki Tydzień') kod = 'WP6';
    else if (nazwa === 'Oktawa Wielkiej Nocy') kod = 'WLK1';
    else if (nazwa === 'Okres Bożego Narodzenia') kod = 'BN';
    return 'I_' + kod + '_' + fmt(mon) + '.docx';
  }

  // Dzieli dni na tygodnie (pon.–niedz.) i buduje akapity każdego tygodnia
  function budujTygodnie(days, opts) {
    var o = {}, k;
    for (k in DEFAULTS) o[k] = DEFAULTS[k];
    for (k in (opts || {})) o[k] = opts[k];
    var weeks = {}, order = [];
    days.forEach(function (d) {
      var mon = addDays(d.date, -((wd(d.date) + 6) % 7));
      if (!weeks[mon]) { weeks[mon] = []; order.push(mon); }
      weeks[mon].push(d);
    });
    order.sort(function (a, b) { return a - b; });
    return order.map(function (mon) {
      var sun = addDays(mon, 6), ps = [];
      var nazwa = nazwaTygodnia(mon);
      ps.push(P(ST_SYG, [{ t: nazwa + ' ', color: '538135' }, { t: '| ' + fmt(mon) + ' – ' + fmt(sun) }]));
      if (o.uwagaTygodnia) ps.push(P(ST_TXT, [{ t: o.uwagaTygodnia, i: true }]));
      weeks[mon].sort(function (a, b) { return a.date - b.date; }).forEach(function (d) { ps = ps.concat(budujDzien(d, o)); });
      return { monday: mon, sunday: sun, nazwa: nazwa, plik: nazwaPliku(nazwa, mon), label: nazwa + ' | ' + fmt(mon) + ' – ' + fmt(sun), paragraphs: ps };
    });
  }

  // ---------- OOXML ----------
  function esc(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  var STYLE_ID = { 'Dzień tygodnia': 'Dzietygodnia', 'Kolekta i sygnaturki': 'Kolektaisygnaturki', 'Bez odstępów1': 'Bezodstpw1' };

  function runXml(r) {
    var pr = '';
    if (r.b) pr += '<w:b/><w:bCs/>';
    if (r.i) pr += '<w:i/><w:iCs/>';
    if (r.color) pr += '<w:color w:val="' + r.color + '"/>';
    if (r.hl) pr += '<w:highlight w:val="' + r.hl + '"/>';
    pr = pr ? '<w:rPr>' + pr + '</w:rPr>' : '';
    return r.t.split('\t').map(function (part, i) {
      return '<w:r>' + pr + (i ? '<w:tab/>' : '') + (part ? '<w:t xml:space="preserve">' + esc(part) + '</w:t>' : '') + '</w:r>';
    }).join('');
  }
  function bodyXml(paragraphs) {
    return paragraphs.map(function (p) {
      var sp = p.style === ST_TXT ? '<w:spacing w:line="240" w:lineRule="auto"/>' : '';
      return '<w:p><w:pPr><w:pStyle w:val="' + STYLE_ID[p.style] + '"/>' + sp + '</w:pPr>' + p.runs.map(runXml).join('') + '</w:p>';
    }).join('');
  }

  // Definicje stylów z szablonu (spłaszczone, żeby działały też w pustym dokumencie).
  // Jeśli dokument ma już te style, Word zostawia swoje.
  var STYLES_XML =
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Dzietygodnia"><w:name w:val="Dzień tygodnia"/><w:qFormat/>' +
    '<w:pPr><w:tabs><w:tab w:val="left" w:pos="142"/><w:tab w:val="left" w:pos="567"/></w:tabs><w:spacing w:before="240" w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:left="851" w:hanging="851"/><w:jc w:val="both"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Metropolis" w:hAnsi="Metropolis" w:cs="Times New Roman"/><w:b/><w:caps/><w:sz w:val="24"/><w:szCs w:val="23"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Kolektaisygnaturki"><w:name w:val="Kolekta i sygnaturki"/><w:qFormat/>' +
    '<w:pPr><w:tabs><w:tab w:val="left" w:pos="142"/><w:tab w:val="left" w:pos="567"/></w:tabs><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Metropolis" w:hAnsi="Metropolis" w:cs="Times New Roman"/><w:caps/><w:sz w:val="18"/><w:szCs w:val="23"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Bezodstpw1"><w:name w:val="Bez odstępów1"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/><w:ind w:left="708" w:hanging="708"/><w:jc w:val="both"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Open Sans" w:eastAsia="Calibri" w:hAnsi="Open Sans" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="24"/></w:rPr></w:style>';

  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  function toFlatOpc(paragraphs) {
    var CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'styles" Target="styles.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="' + CT + 'document.main+xml"><pkg:xmlData>' +
      '<w:document ' + NS + '><w:body>' + bodyXml(paragraphs) + '</w:body></w:document>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="' + CT + 'styles+xml"><pkg:xmlData>' +
      '<w:styles ' + NS + '>' + STYLES_XML + '</w:styles>' +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }

  // ---------- wejście z pliku ----------
  // JSZip i DOMParser podawane z zewnątrz (przeglądarka: globalne; Node: moduły)
  function readFaraDocx(data, JSZip, DOMParserImpl) {
    return JSZip.loadAsync(data).then(function (zip) {
      var f = zip.file('word/document.xml');
      if (!f) throw new Error('Plik nie jest dokumentem Worda (.docx).');
      return f.async('string');
    }).then(function (xml) {
      return parseFaraXml(new DOMParserImpl().parseFromString(xml, 'application/xml'));
    });
  }

  var api = {
    DEFAULTS: DEFAULTS, nazwaTygodnia: nazwaTygodnia, parseFaraXml: parseFaraXml, readFaraDocx: readFaraDocx,
    budujTygodnie: budujTygodnie, bodyXml: bodyXml, toFlatOpc: toFlatOpc, czyscIntencje: czyscIntencje, STYLES_XML: STYLES_XML
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Fara = api;
})(this);
