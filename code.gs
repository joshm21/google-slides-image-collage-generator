// all units in points
// 1 inch = 2.54 cm = 72 points = 96 pixels

// current settings for portrait 8.5"x11" page
// with 2 rows and  2 columns of images
// with 0.5" border around each image
const PAGE_WIDTH = 612
const PAGE_HEIGHT = 792
const IMAGE_BORDER = 36
const IMAGE_ROWS = 2
const IMAGE_COLS = 2

const IMAGES = [
  // [url, orientation(portrait/landscape)]
]

/**
 * @OnlyCurrentDoc
 */

const generate = () => {

  const calcPositions = () => {
    let positions = []
    for (let row = 0; row < IMAGE_ROWS; row++) {
      for (let col = 0; col < IMAGE_COLS; col++) {
        const left = IMAGE_BORDER + (col * cellWidth)
        const top = IMAGE_BORDER + (row * cellHeight)
        const centerX = cellWidth / 2 + (col * cellWidth)
        const centerY = cellHeight / 2 + (row * cellHeight)
        positions.push({ left, top, centerX, centerY })
      }
    }
    return positions
  }

  const setCenter = (image, x, y) => {
    const currentLeft = image.getLeft()
    const currentTop = image.getTop()
    const currentCenterX = currentLeft + (image.getWidth() / 2)
    const currentCenterY = currentTop + (image.getHeight() / 2)
    const shiftLeft = x - currentCenterX
    const shiftTop = y - currentCenterY
    image.setLeft(currentLeft + shiftLeft)
    image.setTop(currentTop + shiftTop)
  }

  const presentation = SlidesApp.getActivePresentation()

  const cellHeight = PAGE_HEIGHT / IMAGE_ROWS
  const cellWidth = PAGE_WIDTH / IMAGE_COLS

  const maxImageHeight = cellHeight - (2 * IMAGE_BORDER)
  const maxImageWidth = cellWidth - (2 * IMAGE_BORDER)

  const imagesPerPage = IMAGE_ROWS * IMAGE_COLS
  const positions = calcPositions()

  let imageIndex = 0
  const numPages = Math.ceil(IMAGES.length / imagesPerPage)
  for (let page = 0; page < numPages; page) {
    const slide = presentation.appendSlide()

    for (let position of positions) {
      const url = IMAGES[imageIndex][0]
      const orientation = IMAGES[imageIndex][1]

      if (orientation == "portrait") {
        slide.insertImage(url, position.left, position.top, maxImageWidth, maxImageHeight)

      } else {
        const image = slide.insertImage(url, 0, 0, maxImageHeight, maxImageWidth).setRotation(90)
        setCenter(image, position.centerX, position.centerY)
      }

      imageIndex += 1
      console.log("Inserted image #" + imageIndex + " of " + IMAGES.length)
      if (imageIndex == IMAGES.length) { return }
    }
  }
}

// Images must be less than 50MB in size, cannot exceed 25 megapixels, and must be in either in PNG, JPEG, or GIF format
// URL must be no larger than 2kB

// can only change page size using advanced services api and creating new presentation :(
// https://developers.google.com/slides/reference/rest/v1/Size
