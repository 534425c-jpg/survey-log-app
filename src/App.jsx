import React, { useState } from "react";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
} from "docx";

const emptyForm = {
  date: new Date().toISOString().slice(0, 10),
  weather: "맑음",
  siteName: "",
  team: "",
  contents: "",
  labor: "",
  equipment: "",
  note: "",
  gps: "",
  photos: [],
};

function koreanDate(dateString) {
  const d = new Date(dateString);
  const days = ["일", "월", "화", "수", "목", "금", "토"];
  return `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}. (${days[d.getDay()]})`;
}

export default function App() {
  const [form, setForm] = useState(emptyForm);

  const update = (key, value) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const getGps = () =>
    new Promise((resolve) => {
      if (!navigator.geolocation) {
        alert("이 기기에서는 위치 기능을 사용할 수 없습니다.");
        return resolve("");
      }

      navigator.geolocation.getCurrentPosition(
        (pos) => {
          const lat = pos.coords.latitude.toFixed(6);
          const lng = pos.coords.longitude.toFixed(6);
          resolve(`N ${lat}, E ${lng}`);
        },
        () => {
          alert("위치 권한을 허용해야 사진 좌표가 저장됩니다.");
          resolve("");
        },
        {
          enableHighAccuracy: true,
          timeout: 10000,
          maximumAge: 0,
        }
      );
    });

  const getCurrentGps = async () => {
    const gps = await getGps();
    if (gps) update("gps", gps);
  };

  const handlePhotos = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    const gps = await getGps();
    const time = new Date().toLocaleString("ko-KR");

    const photos = await Promise.all(
      files.map(async (file) => ({
        name: file.name,
        url: URL.createObjectURL(file),
        dataUrl: await fileToDataUrl(file),
        gps,
        time,
      }))
    );

    setForm((prev) => ({
      ...prev,
      gps: prev.gps || gps,
      photos: [...prev.photos, ...photos],
    }));
  };

  const removePhoto = (index) => {
    setForm((prev) => ({
      ...prev,
      photos: prev.photos.filter((_, i) => i !== index),
    }));
  };

  const downloadDocx = async () => {
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "유적조사일지",
                  bold: true,
                  size: 36,
                }),
              ],
            }),
            new Paragraph(""),
            new Paragraph(`일자: ${koreanDate(form.date)}`),
            new Paragraph(`날씨: ${form.weather}`),
            new Paragraph(`유적명: ${form.siteName}`),
            new Paragraph(`조사단: ${form.team}`),
            new Paragraph(`GPS: ${form.gps}`),
            new Paragraph(""),
            new Paragraph({
              children: [new TextRun({ text: "조사내용", bold: true })],
            }),
            new Paragraph(form.contents || ""),
            new Paragraph(""),
            new Paragraph(`인부: ${form.labor}`),
            new Paragraph(`장비: ${form.equipment}`),
            new Paragraph(`기타사항: ${form.note}`),
            new Paragraph(""),
            ...form.photos.flatMap((p, i) => [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `사진 ${i + 1}. ${p.name}`,
                    bold: true,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new ImageRun({
                    data: dataUrlToUint8Array(p.dataUrl),
                    transformation: {
                      width: 500,
                      height: 350,
                    },
                  }),
                ],
              }),
              new Paragraph(`촬영시각: ${p.time}`),
              new Paragraph(`좌표: ${p.gps || "위치 정보 없음"}`),
              new Paragraph(""),
            ]),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, `${form.siteName || "유적조사일지"}.docx`);
  };

  const downloadPdf = () => {
    window.print();
  };

  return (
    <div style={styles.page}>
      <style>{printStyle}</style>

      <div style={styles.container}>
        <h1 style={styles.title}>유적조사일지</h1>

        <section style={styles.card}>
          <div style={styles.grid3}>
            <Field label="일자">
              <input
                style={styles.input}
                type="date"
                value={form.date}
                onChange={(e) => update("date", e.target.value)}
              />
            </Field>

            <Field label="날씨">
              <input
                style={styles.input}
                value={form.weather}
                onChange={(e) => update("weather", e.target.value)}
              />
            </Field>

            <Field label="GPS 좌표">
              <div style={styles.row}>
                <input
                  style={styles.input}
                  value={form.gps}
                  onChange={(e) => update("gps", e.target.value)}
                  placeholder="N 35..., E 126..."
                />
                <button style={styles.smallButton} onClick={getCurrentGps}>
                  위치
                </button>
              </div>
            </Field>
          </div>

          <Field label="유적명">
            <input
              style={styles.input}
              value={form.siteName}
              onChange={(e) => update("siteName", e.target.value)}
              placeholder="유적명 입력"
            />
          </Field>

          <Field label="조사단">
            <input
              style={styles.input}
              value={form.team}
              onChange={(e) => update("team", e.target.value)}
              placeholder="예: 박철원, 오정훈, 한정훈, 공종찬, 오세찬"
            />
          </Field>

          <Field label="조사내용">
            <textarea
              style={styles.textarea}
              value={form.contents}
              onChange={(e) => update("contents", e.target.value)}
              placeholder="조사내용 입력"
            />
          </Field>

          <div style={styles.grid3}>
            <Field label="인부">
              <input
                style={styles.input}
                value={form.labor}
                onChange={(e) => update("labor", e.target.value)}
              />
            </Field>

            <Field label="장비">
              <input
                style={styles.input}
                value={form.equipment}
                onChange={(e) => update("equipment", e.target.value)}
              />
            </Field>

            <Field label="기타사항">
              <input
                style={styles.input}
                value={form.note}
                onChange={(e) => update("note", e.target.value)}
              />
            </Field>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.sectionTitle}>사진 첨부</h2>
          <p style={styles.helpText}>
            사진을 찍거나 첨부하면 현재 GPS 좌표와 시간이 사진별로 저장됩니다.
          </p>

          <input
            type="file"
            accept="image/*"
            capture="environment"
            multiple
            onChange={handlePhotos}
          />

          <div style={styles.photoGrid}>
            {form.photos.map((p, i) => (
              <div key={i} style={styles.photoBox}>
                <img src={p.url} alt={p.name} style={styles.photo} />
                <div style={styles.photoText}>{p.name}</div>
                <div style={styles.photoText}>촬영: {p.time}</div>
                <div style={styles.photoText}>
                  좌표: {p.gps || "위치 정보 없음"}
                </div>
                <button style={styles.deleteButton} onClick={() => removePhoto(i)}>
                  삭제
                </button>
              </div>
            ))}
          </div>
        </section>

        <section style={styles.card}>
          <div style={styles.buttonRow}>
            <button style={styles.mainButton} onClick={downloadDocx}>
              DOCX 출력
            </button>
            <button style={styles.mainButton} onClick={downloadPdf}>
              PDF 출력
            </button>
          </div>

          <h2 style={styles.sectionTitle}>미리보기</h2>
          <pre style={styles.preview}>
{`일자: ${koreanDate(form.date)}
날씨: ${form.weather}
유적명: ${form.siteName}
조사단: ${form.team}
GPS: ${form.gps}

조사내용
${form.contents}

인부: ${form.labor}
장비: ${form.equipment}
기타사항: ${form.note}

사진 수: ${form.photos.length}장`}
          </pre>
        </section>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label style={styles.field}>
      <span style={styles.label}>{label}</span>
      {children}
    </label>
  );
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

const styles = {
  page: {
    minHeight: "100vh",
    background: "#111827",
    padding: "30px 12px",
    color: "#111827",
    fontFamily: "Arial, sans-serif",
  },
  container: {
    maxWidth: 920,
    margin: "0 auto",
  },
  title: {
    color: "white",
    textAlign: "center",
    fontSize: 42,
    marginBottom: 24,
  },
  card: {
    background: "white",
    borderRadius: 16,
    padding: 22,
    marginBottom: 20,
    boxShadow: "0 8px 24px rgba(0,0,0,0.18)",
  },
  sectionTitle: {
    marginTop: 0,
    marginBottom: 8,
    fontSize: 22,
  },
  helpText: {
    color: "#64748b",
    fontSize: 14,
    marginBottom: 14,
  },
  grid3: {
    display: "grid",
    gridTemplateColumns: "repeat(3, 1fr)",
    gap: 14,
  },
  field: {
    display: "flex",
    flexDirection: "column",
    marginBottom: 14,
    gap: 6,
  },
  label: {
    fontWeight: "bold",
    color: "#64748b",
    fontSize: 14,
  },
  input: {
    width: "100%",
    boxSizing: "border-box",
    padding: "10px 12px",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    fontSize: 15,
    background: "#f8fafc",
    color: "#111827",
  },
  textarea: {
    width: "100%",
    boxSizing: "border-box",
    minHeight: 150,
    padding: "12px",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    fontSize: 15,
    resize: "vertical",
    background: "#f8fafc",
    color: "#111827",
  },
  row: {
    display: "flex",
    gap: 8,
  },
  smallButton: {
    padding: "8px 12px",
    border: "0",
    borderRadius: 8,
    background: "#475569",
    color: "white",
    fontWeight: "bold",
    cursor: "pointer",
  },
  photoGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
    gap: 14,
    marginTop: 16,
  },
  photoBox: {
    border: "1px solid #e5e7eb",
    borderRadius: 12,
    overflow: "hidden",
    background: "#f8fafc",
    paddingBottom: 10,
  },
  photo: {
    width: "100%",
    height: 140,
    objectFit: "cover",
    display: "block",
  },
  photoText: {
    fontSize: 12,
    color: "#475569",
    padding: "4px 8px",
    wordBreak: "break-all",
  },
  deleteButton: {
    marginLeft: 8,
    marginTop: 4,
    padding: "6px 10px",
    border: "0",
    borderRadius: 6,
    background: "#ef4444",
    color: "white",
    cursor: "pointer",
  },
  buttonRow: {
    display: "flex",
    gap: 10,
    justifyContent: "center",
    marginBottom: 18,
  },
  mainButton: {
    padding: "12px 22px",
    border: "0",
    borderRadius: 10,
    background: "#374151",
    color: "white",
    fontSize: 16,
    fontWeight: "bold",
    cursor: "pointer",
  },
  preview: {
    background: "#f1f5f9",
    borderRadius: 12,
    padding: 18,
    whiteSpace: "pre-wrap",
    lineHeight: 1.7,
    color: "#334155",
  },
};

const printStyle = `
@media print {
  body {
    background: white !important;
  }
  button, input[type="file"] {
    display: none !important;
  }
}
@media (max-width: 760px) {
  h1 {
    font-size: 32px !important;
  }
  div {
    max-width: 100%;
  }
}
`;