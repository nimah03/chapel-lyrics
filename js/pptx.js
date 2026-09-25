/* 교육채플 가사 PPT — 원본과 같은 16:9, 검정 배경, G마켓 산스, 두 줄 간격 */
(function (global) {
  const FONT_MEDIUM = "G마켓 산스 TTF Medium";
  const FONT_BOLD = "G마켓 산스 TTF Bold";

  function lyricFontSize(lines) {
    const longest = lines.reduce((max, line) => Math.max(max, Array.from(line || "").length), 0);
    if (longest <= 18) return 46;
    if (longest <= 22) return 40;
    if (longest <= 26) return 34;
    if (longest <= 32) return 30;
    return 26;
  }

  function pairLines(text) {
    const lines = String(text || "")
      .replace(/\r\n/g, "\n")
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
    const pairs = [];
    for (let i = 0; i < lines.length; i += 2) {
      pairs.push({
        line1: lines[i],
        line2: lines[i + 1] || "",
      });
    }
    return pairs;
  }

  function slideKey(slide) {
    if (slide.kind === "title") return `title\n${slide.title}`;
    return `lyric\n${slide.line1}\n${slide.line2}`;
  }

  function buildSlides(songs) {
    const slides = [];
    const seen = new Set();
    const pushUnique = (slide) => {
      const key = slideKey(slide);
      if (seen.has(key)) return;
      seen.add(key);
      slides.push(slide);
    };
    songs.forEach((song) => {
      const title = (song.title || "").trim();
      const pairs = pairLines(song.lyrics);
      if (!title && pairs.length === 0) return;
      if (title) pushUnique({ kind: "title", title });
      pairs.forEach((pair) => {
        pushUnique({
          kind: "lyric",
          line1: pair.line1,
          line2: pair.line2,
          size: lyricFontSize([pair.line1, pair.line2].filter(Boolean)),
        });
      });
    });
    return slides;
  }

  function addLyricText(slide, line1, line2, size) {
    const runs = [{ text: line1, options: { breakLine: true } }];
    if (line2) {
      runs.push({ text: "", options: { breakLine: true } });
      runs.push({ text: line2 });
    }
    slide.addText(runs, {
      x: 0,
      y: 2.538,
      w: 13.333,
      h: 2.424,
      fontFace: FONT_MEDIUM,
      fontSize: size,
      color: "FFFFFF",
      align: "center",
      valign: "top",
      lang: "ko-KR",
      margin: 0,
      isTextBox: true,
    });
  }

  async function buildPptxBlob(songs) {
    const Pptx = global.PptxGenJS;
    if (!Pptx) throw new Error("PPT 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.");
    const slides = buildSlides(songs);
    if (slides.length === 0) throw new Error("곡 제목이나 가사를 입력해 주세요.");

    const pptx = new Pptx();
    pptx.defineLayout({ name: "CHAPEL", width: 13.333, height: 7.5 });
    pptx.layout = "CHAPEL";
    pptx.author = "교육채플 가사";
    pptx.title = "교육채플 가사";

    slides.forEach((slide) => {
      const page = pptx.addSlide();
      page.background = { color: "171717" };
      if (slide.kind === "title") {
        page.addText(slide.title, {
          x: 0,
          y: 3.076,
          w: 13.333,
          h: 1.363,
          fontFace: FONT_BOLD,
          fontSize: 75,
          color: "FFFFFF",
          align: "center",
          valign: "middle",
          bold: true,
          lang: "ko-KR",
          margin: 0,
          isTextBox: true,
        });
      }
      if (slide.kind === "lyric") addLyricText(page, slide.line1, slide.line2, slide.size);
    });

    const inBrowser = typeof window !== "undefined" && typeof document !== "undefined";
    const out = await pptx.write({ outputType: inBrowser ? "blob" : "nodebuffer" });
    if (!inBrowser) return out;
    if (out instanceof Blob) return out;
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  }

  global.ChapelPpt = {
    pairLines,
    buildSlides,
    buildPptxBlob,
    FONT_MEDIUM,
    FONT_BOLD,
  };
})(typeof window !== "undefined" ? window : global);
