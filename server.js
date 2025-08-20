import express from 'express';
import cors from 'cors';
import PptxGenJS from 'pptxgenjs';
import { JSDOM } from 'jsdom';

const app = express();
const PORT = process.env.PORT || 8000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'HTML to PPTX API is running' });
});

// Main conversion endpoint
app.post('/api/convert', async (req, res) => {
  try {
    const { slides } = req.body;
    
    if (!slides || !Array.isArray(slides)) {
      return res.status(400).json({ error: 'Invalid slides data' });
    }

    const pptx = new PptxGenJS();
    
    // Set presentation properties
    pptx.author = 'HTML to PPTX Converter';
    pptx.company = 'Demo Company';
    pptx.revision = '1.0.0';
    pptx.subject = 'Converted Presentation';
    pptx.title = 'Presentation';
    
    // Define master slide layout
    pptx.defineSlideMaster({
      title: 'MASTER_SLIDE',
      background: { color: 'FFFFFF' },
      margin: [0.5, 0.5, 0.5, 0.5],
    });

    // Process each slide
    for (const [index, slideData] of slides.entries()) {
      const slide = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
      
      // Handle different layouts
      switch(slideData.layout) {
        case 'title':
          // Title slide layout
          if (slideData.title) {
            slide.addText(slideData.title, {
              x: 1,
              y: 2.5,
              w: 8,
              h: 1.5,
              fontSize: 40,
              bold: true,
              color: slideData.titleColor || '2D3748',
              align: slideData.centerAlign ? 'center' : 'left',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Add subtitle if present
          if (slideData.subtitle) {
            slide.addText(slideData.subtitle, {
              x: 1,
              y: 4,
              w: 8,
              h: 1,
              fontSize: 24,
              color: slideData.subtitleColor || '4A5568',
              align: slideData.centerAlign ? 'center' : 'left',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          break;
          
        case 'twoColumn':
          // Title with underline for two column layout
          if (slideData.title) {
            slide.addText(slideData.title, {
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 0.8,
              fontSize: 28,
              bold: true,
              color: slideData.titleColor || '2D3748',
              align: 'left',
              fontFace: slideData.fontFamily || 'Calibri'
            });
            
            // Add blue underline
            slide.addShape(pptx.ShapeType.line, {
              x: 0.5,
              y: 1.3,
              w: 9,
              h: 0,
              line: { color: '4472C4', width: 2 }
            });
          }
          
          // Left column
          if (slideData.leftContent) {
            slide.addText(slideData.leftContent, {
              x: 0.5,
              y: 1.8,
              w: 4.2,
              h: 4,
              fontSize: 14,
              color: slideData.contentColor || '4A5568',
              align: 'left',
              valign: 'top',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Right column
          if (slideData.rightContent) {
            slide.addText(slideData.rightContent, {
              x: 5,
              y: 1.8,
              w: 4.2,
              h: 4,
              fontSize: 14,
              color: slideData.contentColor || '4A5568',
              align: 'left',
              valign: 'top',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          break;
          
        case 'imageWithText':
          // Title with underline for image with text layout
          if (slideData.title) {
            slide.addText(slideData.title, {
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 0.8,
              fontSize: 28,
              bold: true,
              color: slideData.titleColor || '2D3748',
              align: 'left',
              fontFace: slideData.fontFamily || 'Calibri'
            });
            
            // Add blue underline
            slide.addShape(pptx.ShapeType.line, {
              x: 0.5,
              y: 1.3,
              w: 9,
              h: 0,
              line: { color: '4472C4', width: 2 }
            });
          }
          
          // Image on the left
          if (slideData.image && slideData.image.url) {
            try {
              const imageOptions = {
                x: 0.5,
                y: 1.8,
                w: 4,
                h: 3.5,
                sizing: { type: 'contain' }
              };

              if (slideData.image.url.startsWith('data:image/')) {
                // For base64 data URLs (any format: png, jpg, jpeg, gif, etc.)
                // Clean and validate base64 data
                const cleanedBase64 = slideData.image.url.trim().replace(/\s/g, '');
                slide.addImage({
                  data: cleanedBase64,
                  ...imageOptions
                });
              } else if (slideData.image.url.startsWith('http://') || slideData.image.url.startsWith('https://')) {
                // For online URLs - fetch and convert to base64 to avoid CORS issues
                try {
                  console.log('Fetching image from URL (imageWithText):', slideData.image.url);
                  const response = await fetch(slideData.image.url);
                  if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                  }
                  const arrayBuffer = await response.arrayBuffer();
                  const buffer = Buffer.from(arrayBuffer);
                  const base64 = buffer.toString('base64');
                  const contentType = response.headers.get('content-type') || 'image/jpeg';
                  const dataUrl = `data:${contentType};base64,${base64}`;
                  
                  slide.addImage({
                    data: dataUrl,
                    ...imageOptions
                  });
                  console.log('Successfully added image from URL (imageWithText)');
                } catch (urlError) {
                  console.log('Failed to fetch image from URL (imageWithText):', urlError.message);
                  // Fallback: try direct path method
                  slide.addImage({
                    path: slideData.image.url,
                    ...imageOptions
                  });
                }
              } else {
                // For local file paths
                slide.addImage({
                  path: slideData.image.url,
                  ...imageOptions
                });
              }
            } catch (err) {
              console.log('Failed to add image:', err.message);
              // Add placeholder rectangle
              slide.addShape(pptx.ShapeType.rect, {
                x: 0.5,
                y: 1.8,
                w: 4,
                h: 3.5,
                fill: { color: 'E0E0E0' },
                line: { color: '999999', width: 1 }
              });
              slide.addText('Image Placeholder', {
                x: 0.5,
                y: 3.3,
                w: 4,
                h: 0.5,
                align: 'center',
                color: '666666',
                fontSize: 14
              });
            }
          }
          
          // Text content on the right
          if (slideData.text) {
            slide.addText(slideData.text, {
              x: 5,
              y: 1.8,
              w: 4.2,
              h: 3.5,
              fontSize: 14,
              color: slideData.contentColor || '4A5568',
              align: 'left',
              valign: 'top',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          break;
          
        case 'comparison':
          // Comparison layout with left and right columns
          if (slideData.title) {
            slide.addText(slideData.title, {
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 0.8,
              fontSize: 28,
              bold: true,
              color: slideData.titleColor || '2D3748',
              align: 'center',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Left column title
          if (slideData.leftTitle) {
            slide.addText(slideData.leftTitle, {
              x: 0.5,
              y: 1.5,
              w: 4.2,
              h: 0.6,
              fontSize: 20,
              bold: true,
              color: '#1F4E79',
              align: 'center',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Right column title
          if (slideData.rightTitle) {
            slide.addText(slideData.rightTitle, {
              x: 5,
              y: 1.5,
              w: 4.2,
              h: 0.6,
              fontSize: 20,
              bold: true,
              color: '#70AD47',
              align: 'center',
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Left column bullets
          if (slideData.leftBullets && slideData.leftBullets.length > 0) {
            const leftBulletOptions = slideData.leftBullets.map(bullet => ({
              text: bullet,
              options: {
                bullet: true,
                fontSize: 16,
                color: '4A5568'
              }
            }));
            
            slide.addText(leftBulletOptions, {
              x: 0.5,
              y: 2.3,
              w: 4.2,
              h: 3.5,
              fontSize: 16,
              bullet: true,
              lineSpacing: 32,
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          
          // Right column bullets
          if (slideData.rightBullets && slideData.rightBullets.length > 0) {
            const rightBulletOptions = slideData.rightBullets.map(bullet => ({
              text: bullet,
              options: {
                bullet: true,
                fontSize: 16,
                color: '4A5568'
              }
            }));
            
            slide.addText(rightBulletOptions, {
              x: 5,
              y: 2.3,
              w: 4.2,
              h: 3.5,
              fontSize: 16,
              bullet: true,
              lineSpacing: 32,
              fontFace: slideData.fontFamily || 'Calibri'
            });
          }
          break;
          
        default:
          // Title and content layout (default)
          if (slideData.title) {
            slide.addText(slideData.title, {
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 0.8,
              fontSize: 28,
              bold: true,
              color: slideData.titleColor || '2D3748',
              align: slideData.centerAlign ? 'center' : 'left',
              fontFace: slideData.fontFamily || 'Calibri'
            });
            
            // Add blue underline for title and content layout
            slide.addShape(pptx.ShapeType.line, {
              x: 0.5,
              y: 1.3,
              w: 9,
              h: 0,
              line: { color: '4472C4', width: 2 }
            });
          }
          break;
      }
      
      // Add bullet points if present (for non-title and non-comparison layouts)  
      if (slideData.layout !== 'title' && slideData.layout !== 'comparison' && slideData.bullets && slideData.bullets.length > 0) {
        const yPosition = slideData.title ? 1.8 : 1;
        
        // Use EXACTLY the same format as comparison leftBullets/rightBullets
        const bulletOptions = slideData.bullets.map(bullet => {
          const bulletText = typeof bullet === 'string' ? bullet : bullet.text;
          return {
            text: bulletText,
            options: {
              bullet: true,
              fontSize: 18,
              color: slideData.bulletColor || '4A5568'
            }
          };
        });
        
        // Use EXACTLY the same addText call as comparison
        slide.addText(bulletOptions, {
          x: 0.5,
          y: yPosition,
          w: 9,
          h: 4.5,
          fontSize: 18,
          bullet: true,
          lineSpacing: 32,
          fontFace: slideData.fontFamily || 'Calibri'
        });
      }
      
      // Add image if present (for other layouts)
      if (slideData.layout !== 'imageWithText' && slideData.image && slideData.image.url) {
        try {
          const imageOptions = {
            x: slideData.image.x || 1,
            y: slideData.image.y || 3,
            w: slideData.image.width || 3,
            h: slideData.image.height || 2,
            sizing: { type: 'contain' }
          };

          if (slideData.image.url.startsWith('data:image/')) {
            // For base64 data URLs (any format: png, jpg, jpeg, gif, etc.)
            // Clean and validate base64 data
            const cleanedBase64 = slideData.image.url.trim().replace(/\s/g, '');
            slide.addImage({
              data: cleanedBase64,
              ...imageOptions
            });
          } else if (slideData.image.url.startsWith('http://') || slideData.image.url.startsWith('https://')) {
            // For online URLs - fetch and convert to base64 to avoid CORS issues
            try {
              console.log('Fetching image from URL:', slideData.image.url);
              const response = await fetch(slideData.image.url);
              if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
              }
              const arrayBuffer = await response.arrayBuffer();
              const buffer = Buffer.from(arrayBuffer);
              const base64 = buffer.toString('base64');
              const contentType = response.headers.get('content-type') || 'image/jpeg';
              const dataUrl = `data:${contentType};base64,${base64}`;
              
              slide.addImage({
                data: dataUrl,
                ...imageOptions
              });
              console.log('Successfully added image from URL');
            } catch (urlError) {
              console.log('Failed to fetch image from URL:', urlError.message);
              // Fallback: try direct path method
              slide.addImage({
                path: slideData.image.url,
                ...imageOptions
              });
            }
          } else {
            // For local file paths
            slide.addImage({
              path: slideData.image.url,
              ...imageOptions
            });
          }
        } catch (err) {
          console.log('Failed to add image:', err.message);
          // Add error placeholder
          slide.addShape(pptx.ShapeType.rect, {
            x: slideData.image.x || 1,
            y: slideData.image.y || 3,
            w: slideData.image.width || 3,
            h: slideData.image.height || 2,
            fill: { color: 'F0F0F0' },
            line: { color: 'CCCCCC', width: 1 }
          });
          slide.addText('Image Error', {
            x: slideData.image.x || 1,
            y: (slideData.image.y || 3) + (slideData.image.height || 2)/2 - 0.2,
            w: slideData.image.width || 3,
            h: 0.4,
            fontSize: 12,
            color: '666666',
            align: 'center'
          });
        }
      }
      
      // Add custom text blocks (only for non-title layouts)
      if (slideData.layout !== 'title' && slideData.textBlocks) {
        slideData.textBlocks.forEach(block => {
          slide.addText(block.text, {
            x: block.x || 0.5,
            y: block.y || 2,
            w: block.width || 9,
            h: block.height || 1,
            fontSize: block.fontSize || 14,
            color: block.color || '333333',
            align: block.align || 'left',
            bold: block.bold || false,
            italic: block.italic || false,
            fontFace: block.fontFamily || 'Arial'
          });
        });
      }
    }

    // Generate PPTX as base64
    const pptxBase64 = await pptx.write({ outputType: 'base64' });
    
    res.json({
      success: true,
      data: pptxBase64,
      filename: `presentation_${Date.now()}.pptx`
    });
    
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ 
      error: 'Failed to convert slides to PPTX',
      details: error.message 
    });
  }
});

// HTML parsing endpoint
app.post('/api/parse-html', async (req, res) => {
  try {
    const { html } = req.body;
    
    if (!html) {
      return res.status(400).json({ error: 'HTML content is required' });
    }

    const dom = new JSDOM(html);
    const document = dom.window.document;
    
    const slides = [];
    const slideElements = document.querySelectorAll('.slide');
    
    slideElements.forEach(slideEl => {
      const slideData = {
        title: '',
        titleStyles: {},
        bullets: [],
        textBlocks: [],
        image: null
      };
      
      // Extract title
      const titleEl = slideEl.querySelector('h1, h2, .slide-title');
      if (titleEl) {
        slideData.title = titleEl.textContent.trim();
        slideData.titleStyles = getComputedStyles(titleEl, dom.window);
      }
      
      // Extract bullet points
      const bullets = slideEl.querySelectorAll('ul li, ol li, .bullet-point');
      bullets.forEach(bulletEl => {
        const styles = getComputedStyles(bulletEl, dom.window);
        slideData.bullets.push({
          text: bulletEl.textContent.trim(),
          fontSize: parseInt(styles.fontSize) || 18,
          color: rgbToHex(styles.color) || '4A5568',
          indentLevel: 0
        });
      });
      
      // Extract images
      const imgEl = slideEl.querySelector('img');
      if (imgEl) {
        slideData.image = {
          url: imgEl.src,
          width: 4,
          height: 3,
          x: 3,
          y: 3
        };
      }
      
      // Extract other text blocks
      const textEls = slideEl.querySelectorAll('p, .text-block');
      textEls.forEach(textEl => {
        const styles = getComputedStyles(textEl, dom.window);
        slideData.textBlocks.push({
          text: textEl.textContent.trim(),
          fontSize: parseInt(styles.fontSize) || 14,
          color: rgbToHex(styles.color) || '333333',
          fontFamily: styles.fontFamily || 'Arial',
          bold: styles.fontWeight === 'bold' || parseInt(styles.fontWeight) >= 700,
          italic: styles.fontStyle === 'italic'
        });
      });
      
      slides.push(slideData);
    });
    
    res.json({ success: true, slides });
    
  } catch (error) {
    console.error('HTML parsing error:', error);
    res.status(500).json({ 
      error: 'Failed to parse HTML',
      details: error.message 
    });
  }
});

// Style extraction helper
function extractStyles(styleObj) {
  if (!styleObj) return {};
  
  return {
    fontSize: styleObj.fontSize ? parseInt(styleObj.fontSize) : 18,
    color: styleObj.color ? styleObj.color.replace('#', '') : '333333',
    fontFamily: styleObj.fontFamily || 'Arial',
    bold: styleObj.fontWeight === 'bold' || (styleObj.fontWeight && parseInt(styleObj.fontWeight) >= 700),
    italic: styleObj.fontStyle === 'italic',
    align: styleObj.textAlign || 'left'
  };
}

// Get computed styles from DOM element
function getComputedStyles(element, window) {
  const styles = window.getComputedStyle(element);
  return {
    fontSize: styles.fontSize,
    color: styles.color,
    fontFamily: styles.fontFamily,
    fontWeight: styles.fontWeight,
    fontStyle: styles.fontStyle,
    textAlign: styles.textAlign
  };
}

// Convert RGB to Hex color
function rgbToHex(rgb) {
  if (!rgb) return null;
  
  const match = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (!match) return rgb.replace('#', '');
  
  const hex = (x) => {
    const hexValue = parseInt(x).toString(16);
    return hexValue.length === 1 ? '0' + hexValue : hexValue;
  };
  
  return hex(match[1]) + hex(match[2]) + hex(match[3]);
}

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`API endpoints:`);
  console.log(`  - GET  /api/health`);
  console.log(`  - POST /api/convert`);
  console.log(`  - POST /api/parse-html`);
});