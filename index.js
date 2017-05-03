const fs = require('fs')
const path = require('path')
const AdmZip = require('adm-zip')
const archiver = require('archiver')

// Возвращает список docx-файлов в исходной папке
function getDocs () {
  return docList = fs.readdirSync('input').filter(file => path.extname(file) === '.docx')
}

// заменяет текст в колонтитуле docx-файлв
function docUnpacker (fileName, targetText, newText) {
  let srcPath = path.join(path.join(__dirname, 'input'), fileName)
  let destPath = path.join(path.join(__dirname, 'output'), fileName)
  let headPath = 'word/header1.xml'
  let zip = new AdmZip(srcPath)

  if (zip.getEntry(headPath)) {
    let content = zip.readAsText(headPath)
    if (content.includes(targetText)) {
      content = content.replace(targetText, newText)
      zip.deleteFile(headPath)

      let entries = zip.getEntries()
      let archive = archiver('zip')

      for (let i = 0, n = entries.length; i < n; i++) {
        if (!entries[i].isDirectory) {
            let data = entries[i].getData()
            archive.append(data, {name: entries[i].entryName})
        }
      }
      archive.append(new Buffer(content), {name: headPath})
      let out = fs.createWriteStream(destPath)
      archive.pipe(out)
      archive.finalize()
    } else {
      console.log(`${fileName}: в колонтитуле не найден заменяемый текст`)
    }
  } else {
    console.log(`${fileName}: колонтитул не найден`)
  }
}

let src = '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>Fr</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="de-DE"/></w:rPr><w:t>ü</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="de-DE"/></w:rPr><w:t>hlingstrimester</w:t></w:r>'

let dest = '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>Herbstsemester</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="de-DE"/></w:rPr><w:t></w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl w:val="0"/><w:lang w:val="de-DE"/></w:rPr><w:t></w:t></w:r>'

let files = getDocs()
for (let i = 0; i < files.length; i++) {
  docUnpacker(files[i], src, dest)
}
