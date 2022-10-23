const template = `<!DOCTYPE html>
<html lang="zh" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">

<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>__title__</title>
  <!--[if gte mso 9]><xml> <w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
  <style type="text/css">
    @page Section0 {__option__}
    div.Section0 { page: Section0; }
  </style>
</head>

<body>
  <div class="Section0">__content__</div>
</body>
</html>`


class HTML2Doc{
  /**
   * @constructor HTML2Doc的构造函数
   * @param {string} html html代码
   * @param {object} option 配置项
   * @param {string} option.marginTop 页边距-上（单位是pt）
   * @param {string} option.marginBottom 页边距-下（单位是pt）
   * @param {string} option.marginLeft 页边距-左（单位是pt）
   * @param {string} option.marginRight 页边距-右（单位是pt）
   * @param {string} option.header 页眉的id(不带#)
   * @param {string} option.footer 页脚的id(不带#)
   */
  constructor(html, option = {}) {
    if(option.marginTop === undefined) {
      option.marginTop = HTML2Doc.mm2pt(34.58)
    }
    if(option.marginBottom === undefined) {
      option.marginBottom = HTML2Doc.mm2pt(32.58)
    }
    if(option.marginLeft === undefined) {
      option.marginLeft = HTML2Doc.mm2pt(28)
    }
    if(option.marginRight === undefined) {
      option.marginRight = HTML2Doc.mm2pt(26)
    }
    if(option.header === undefined) {
      option.header= `header`
    }
    if(option.footer === undefined) {
      option.footer= `footer`
    }
    let section0 = `margin-top: ${option.marginTop}pt;
    margin-bottom: ${option.marginBottom}pt;
    margin-left: ${option.marginLeft}pt;
    margin-right: ${option.marginRight}pt;
    mso-header: ${option.header};
    mso-footer: ${option.footer};`
    this.content = template.replace('__option__', section0).replace(/\n/g, '').replace('__content__', html)
  }

  /**
   * 保存word文件
   * @param {string} fileName 保存的文件名
   */
  save(fileName) {
    let link = document.createElement('a')
    this.content = this.content.replace('__title__', fileName)
    let blobUrl = URL.createObjectURL(new Blob([this.content], {
      type: 'application/msword'
    }))
    link.href = blobUrl
    link.download = fileName
    link.click()
    URL.revokeObjectURL(blobUrl)
  }

  /**
   * @static img标签转base64
   * @param {HTMLImageElement} img img标签
   * @returns {string} base64 图片的base64字符串
   */
  static img2b64(img) {
    let imgWidth = img.clientWidth
    let imgHeight = img.clientHeight
    let canvas = document.createElement('canvas')
    canvas.width = imgWidth
    canvas.height = imgHeight
    let canvasCtx = canvas.getContext('2d')
    canvasCtx.drawImage(img, imgWidth, imgHeight)
    return 'data:image/png;base64,'+canvas.toDataURL()
  }

  /**
   * @static 毫米换算pt
   * @param {number} mm 毫米
   * @returns {string} pt 保留4位小数
   */
  static mm2pt(mm) {
    return (mm/0.35278).toFixed(4)
  }
}