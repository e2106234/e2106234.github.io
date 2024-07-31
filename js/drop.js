const dropArea = document.getElementById('drop-area');
const output = document.getElementById('output');

dropArea.addEventListener('dragover', (event) => {
    event.preventDefault();
});

dropArea.addEventListener('dragleave', () => {
});

dropArea.addEventListener('drop', (event) => {
    event.preventDefault();

    const file = event.dataTransfer.files[0];
    if (file && file.name.endsWith('.docx')) {
        readFile(file);
    } else {
        output.innerHTML = 'Wordファイル(.docx)をドロップしてください。';
    }
});

function readFile(file) {
    const reader = new FileReader();
    reader.onload = async function(event) {
        const arrayBuffer = event.target.result;
        const zip = new JSZip();
        await zip.loadAsync(arrayBuffer);

        const content = await zip.file("word/document.xml").async("string");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(content, "application/xml");
        const textNodes = xmlDoc.getElementsByTagName("w:t");
        let text = "";

        for (let i = 0; i < textNodes.length; i++) {
            text += textNodes[i].textContent;
        }

        const charCount = text.length;
        output.innerHTML = `<div class="output-s"><p class="result">文字数: ${charCount}</p><p>${text}</p></div>`;
    };
    reader.readAsArrayBuffer(file);
}