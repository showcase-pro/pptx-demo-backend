import PptxGenJS from 'pptxgenjs';

export class ConverterService {
  constructor() {
    this.layoutTemplates = {
      title: {
        name: 'Title Slide',
        apply: (slide, data) => this.applyTitleLayout(slide, data)
      },
      titleContent: {
        name: 'Title and Content',
        apply: (slide, data) => this.applyTitleContentLayout(slide, data)
      },
      twoColumn: {
        name: 'Two Column',
        apply: (slide, data) => this.applyTwoColumnLayout(slide, data)
      },
      comparison: {
        name: 'Comparison',
        apply: (slide, data) => this.applyComparisonLayout(slide, data)
      },
      imageWithText: {
        name: 'Image with Text',
        apply: (slide, data) => this.applyImageWithTextLayout(slide, data)
      }
    };
  }

  applyTitleLayout(slide, data) {
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 2.5,
        w: 9,
        h: 2,
        fontSize: 44,
        bold: true,
        color: '2D3748',
        align: 'center',
        valign: 'middle'
      });
    }
    
    if (data.subtitle) {
      slide.addText(data.subtitle, {
        x: 1,
        y: 4.5,
        w: 8,
        h: 1,
        fontSize: 24,
        color: '718096',
        align: 'center'
      });
    }
  }

  applyTitleContentLayout(slide, data) {
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 1,
        fontSize: 32,
        bold: true,
        color: '2D3748'
      });
    }
    
    if (data.bullets && data.bullets.length > 0) {
      const bulletTexts = data.bullets.map(bullet => ({
        text: bullet.text,
        options: {
          bullet: true,
          fontSize: bullet.fontSize || 18,
          indentLevel: bullet.indentLevel || 0
        }
      }));
      
      slide.addText(bulletTexts, {
        x: 0.5,
        y: 1.8,
        w: 9,
        h: 4.5,
        fontSize: 18,
        bullet: true,
        lineSpacing: 32
      });
    }
  }

  applyTwoColumnLayout(slide, data) {
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 0.8,
        fontSize: 28,
        bold: true,
        color: '2D3748'
      });
    }
    
    // Left column
    if (data.leftContent) {
      slide.addText(data.leftContent, {
        x: 0.5,
        y: 1.5,
        w: 4,
        h: 4.5,
        fontSize: 16,
        valign: 'top'
      });
    }
    
    // Right column
    if (data.rightContent) {
      slide.addText(data.rightContent, {
        x: 5,
        y: 1.5,
        w: 4,
        h: 4.5,
        fontSize: 16,
        valign: 'top'
      });
    }
  }

  applyComparisonLayout(slide, data) {
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.7,
        fontSize: 28,
        bold: true,
        color: '2D3748',
        align: 'center'
      });
    }
    
    // Left side header
    if (data.leftTitle) {
      slide.addText(data.leftTitle, {
        x: 0.5,
        y: 1.2,
        w: 4,
        h: 0.5,
        fontSize: 20,
        bold: true,
        color: '667EEA',
        align: 'center'
      });
    }
    
    // Right side header
    if (data.rightTitle) {
      slide.addText(data.rightTitle, {
        x: 5,
        y: 1.2,
        w: 4,
        h: 0.5,
        fontSize: 20,
        bold: true,
        color: '48BB78',
        align: 'center'
      });
    }
    
    // Left content
    if (data.leftBullets) {
      const leftTexts = data.leftBullets.map(text => ({
        text,
        options: { bullet: true }
      }));
      
      slide.addText(leftTexts, {
        x: 0.5,
        y: 2,
        w: 4,
        h: 3.5,
        fontSize: 16,
        bullet: true
      });
    }
    
    // Right content
    if (data.rightBullets) {
      const rightTexts = data.rightBullets.map(text => ({
        text,
        options: { bullet: true }
      }));
      
      slide.addText(rightTexts, {
        x: 5,
        y: 2,
        w: 4,
        h: 3.5,
        fontSize: 16,
        bullet: true
      });
    }
  }

  applyImageWithTextLayout(slide, data) {
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.7,
        fontSize: 28,
        bold: true,
        color: '2D3748'
      });
    }
    
    // Image on the left
    if (data.image) {
      slide.addImage({
        path: data.image.url,
        x: 0.5,
        y: 1.5,
        w: 4,
        h: 4
      });
    }
    
    // Text on the right
    if (data.text) {
      slide.addText(data.text, {
        x: 5,
        y: 1.5,
        w: 4,
        h: 4,
        fontSize: 16,
        valign: 'top'
      });
    }
  }

  convertSlidesToPptx(slides, options = {}) {
    const pptx = new PptxGenJS();
    
    // Set presentation properties
    pptx.author = options.author || 'HTML to PPTX Converter';
    pptx.company = options.company || '';
    pptx.title = options.title || 'Presentation';
    
    // Define master slides
    this.defineMasterSlides(pptx);
    
    // Process each slide
    slides.forEach(slideData => {
      const slide = pptx.addSlide({ masterName: slideData.masterName || 'MASTER_SLIDE' });
      
      // Apply layout template if specified
      if (slideData.layout && this.layoutTemplates[slideData.layout]) {
        this.layoutTemplates[slideData.layout].apply(slide, slideData);
      } else {
        // Default layout processing
        this.applyDefaultLayout(slide, slideData);
      }
    });
    
    return pptx;
  }

  applyDefaultLayout(slide, data) {
    // Title
    if (data.title) {
      slide.addText(data.title, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 1,
        fontSize: data.titleFontSize || 32,
        bold: true,
        color: data.titleColor || '2D3748',
        align: data.titleAlign || 'left'
      });
    }
    
    // Bullets
    if (data.bullets && data.bullets.length > 0) {
      const bulletOptions = data.bullets.map(bullet => ({
        text: bullet.text || bullet,
        options: {
          bullet: true,
          fontSize: bullet.fontSize || 18,
          indentLevel: bullet.indentLevel || 0
        }
      }));
      
      slide.addText(bulletOptions, {
        x: 0.5,
        y: 2,
        w: 9,
        h: 4,
        bullet: true,
        lineSpacing: 36
      });
    }
    
    // Custom text blocks
    if (data.textBlocks) {
      data.textBlocks.forEach(block => {
        slide.addText(block.text, {
          x: block.x || 0.5,
          y: block.y || 2,
          w: block.width || 9,
          h: block.height || 1,
          fontSize: block.fontSize || 14,
          color: block.color || '333333',
          align: block.align || 'left',
          bold: block.bold || false,
          italic: block.italic || false
        });
      });
    }
    
    // Images
    if (data.image) {
      slide.addImage({
        path: data.image.url,
        x: data.image.x || 1,
        y: data.image.y || 3,
        w: data.image.width || 3,
        h: data.image.height || 2
      });
    }
  }

  defineMasterSlides(pptx) {
    // Default master
    pptx.defineSlideMaster({
      title: 'MASTER_SLIDE',
      background: { color: 'FFFFFF' },
      margin: [0.5, 0.5, 0.5, 0.5]
    });
    
    // Dark theme master
    pptx.defineSlideMaster({
      title: 'DARK_MASTER',
      background: { color: '1A202C' },
      margin: [0.5, 0.5, 0.5, 0.5]
    });
    
    // Gradient master
    pptx.defineSlideMaster({
      title: 'GRADIENT_MASTER',
      background: {
        fill: {
          type: 'grad',
          colors: [
            { color: '667EEA', position: 0 },
            { color: '764BA2', position: 100 }
          ],
          angle: 135
        }
      },
      margin: [0.5, 0.5, 0.5, 0.5]
    });
  }
}

export default new ConverterService();