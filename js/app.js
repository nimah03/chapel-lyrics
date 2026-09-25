const SAMPLE = {
  title: "나의 영원하신 기업",
  lyrics: `나의 영원하신 기업
생명보다 귀하다
나의 갈 길 다가도록
나와 동행하소서
주님은 나의 힘이요
주님은 나의 방패라
나의 갈 길 다가도록
나와 동행하소서`,
};

const songsEl = document.querySelector("#songs");
const previewEl = document.querySelector("#preview");
const countEl = document.querySelector("#slideCount");
const statusEl = document.querySelector("#status");
const downloadBtn = document.querySelector("#downloadBtn");
const shareBtn = document.querySelector("#shareBtn");
const addSongBtn = document.querySelector("#addSongBtn");

let nextId = 1;
const songs = [];

function setStatus(message, isError) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", Boolean(isError));
}

function readSongs() {
  return songs.map((song) => {
    const card = songsEl.querySelector(`[data-id="${song.id}"]`);
    return {
      title: card.querySelector(".song-title").value,
      lyrics: card.querySelector(".song-lyrics").value,
    };
  });
}

function renderPreview() {
  const slides = ChapelPpt.buildSlides(readSongs());
  previewEl.innerHTML = "";
  if (slides.length === 0) {
    previewEl.innerHTML = `<div class="empty-preview">가사를 입력하면 두 줄씩 나뉜 슬라이드가 여기에 나타납니다.</div>`;
    countEl.textContent = "슬라이드 0장";
    return slides;
  }
  slides.forEach((slide, index) => {
    const card = document.createElement("article");
    card.className = "slide-card";
    const frame = document.createElement("div");
    frame.className = "slide-frame";
    if (slide.kind === "blank") {
      frame.innerHTML = `<p class="slide-kicker">사이</p>`;
    } else if (slide.kind === "title") {
      frame.innerHTML = `<p class="slide-title">${escapeHtml(slide.title)}</p>`;
    } else {
      const size = Math.max(22, Math.round((slide.size / 46) * 34));
      frame.innerHTML =
        `<div class="lyric-block" style="font-size:${size}px">` +
        `<p>${escapeHtml(slide.line1)}</p>` +
        `<p class="lyric-gap"></p>` +
        (slide.line2 ? `<p>${escapeHtml(slide.line2)}</p>` : "") +
        `</div>`;
    }
    const label = document.createElement("p");
    label.className = "slide-label";
    label.textContent =
      slide.kind === "title" ? `${index + 1} · 제목` : slide.kind === "lyric" ? `${index + 1} · 가사` : `${index + 1} · 빈 화면`;
    card.append(frame, label);
    previewEl.append(card);
  });
  const lyricCount = slides.filter((slide) => slide.kind === "lyric").length;
  countEl.textContent = `슬라이드 ${slides.length}장 · 가사 ${lyricCount}장`;
  return slides;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function addSong(data) {
  const id = nextId;
  nextId += 1;
  songs.push({ id });
  const card = document.createElement("section");
  card.className = "song-card";
  card.dataset.id = String(id);
  card.innerHTML = `
    <div class="song-head">
      <h2>곡 ${songs.length}</h2>
      <button type="button" class="text-btn remove-song">삭제</button>
    </div>
    <label>곡 제목
      <input class="song-title" type="text" placeholder="예: 나의 영원하신 기업" value="${escapeHtml(data?.title || "")}">
    </label>
    <label>전체 가사 <span>한 줄에 한 소절. 두 줄씩 한 슬라이드가 되고, 같은 가사는 한 번만 남습니다.</span>
      <textarea class="song-lyrics" rows="10" placeholder="나의 영원하신 기업&#10;생명보다 귀하다&#10;나의 갈 길 다가도록&#10;나와 동행하소서">${escapeHtml(data?.lyrics || "")}</textarea>
    </label>
  `;
  card.querySelector(".remove-song").addEventListener("click", () => {
    const index = songs.findIndex((song) => song.id === id);
    if (songs.length === 1) {
      card.querySelector(".song-title").value = "";
      card.querySelector(".song-lyrics").value = "";
      renderPreview();
      return;
    }
    songs.splice(index, 1);
    card.remove();
    songsEl.querySelectorAll(".song-card h2").forEach((heading, i) => {
      heading.textContent = `곡 ${i + 1}`;
    });
    renderPreview();
  });
  card.addEventListener("input", () => renderPreview());
  songsEl.append(card);
  renderPreview();
}

async function makeFile() {
  const data = readSongs();
  const blob = await ChapelPpt.buildPptxBlob(data);
  const firstTitle = data.map((song) => song.title.trim()).find(Boolean) || "교육채플 가사";
  const safeName = firstTitle.replace(/[\\/:*?"<>|]/g, " ").trim() || "교육채플 가사";
  return { blob, name: `${safeName}.pptx` };
}

downloadBtn.addEventListener("click", async () => {
  try {
    setStatus("PPT 파일을 만들고 있습니다.");
    const { blob, name } = await makeFile();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    link.click();
    URL.revokeObjectURL(url);
    setStatus(`${name} 파일을 내려받았습니다.`);
  } catch (error) {
    setStatus(error.message || "파일을 만들지 못했습니다.", true);
  }
});

shareBtn.addEventListener("click", async () => {
  try {
    setStatus("공유할 PPT 파일을 만들고 있습니다.");
    const { blob, name } = await makeFile();
    const file = new File([blob], name, {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: name });
      setStatus("공유 창을 열었습니다.");
      return;
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    link.click();
    URL.revokeObjectURL(url);
    setStatus("이 브라우저에서는 공유 창을 열 수 없어 파일로 저장했습니다.");
  } catch (error) {
    if (error && error.name === "AbortError") {
      setStatus("공유를 취소했습니다.");
      return;
    }
    setStatus(error.message || "공유하지 못했습니다.", true);
  }
});

addSongBtn.addEventListener("click", () => addSong());
document.querySelector("#sampleBtn").addEventListener("click", () => {
  const first = songsEl.querySelector(".song-card");
  if (!first) addSong(SAMPLE);
  else {
    const title = first.querySelector(".song-title");
    const lyrics = first.querySelector(".song-lyrics");
    if (title.value.trim() || lyrics.value.trim()) addSong(SAMPLE);
    else {
      title.value = SAMPLE.title;
      lyrics.value = SAMPLE.lyrics;
      renderPreview();
    }
  }
});

addSong();
renderPreview();
