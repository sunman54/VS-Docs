import JSZip from 'jszip';

export interface ExtractedImage {
  name: string;
  base64: string;
}

export async function extractImagesFromDocx(buffer: ArrayBuffer): Promise<ExtractedImage[]> {
  const zip = await JSZip.loadAsync(buffer);
  const images: ExtractedImage[] = [];

  const mediaFiles = Object.keys(zip.files).filter(
    (fileName) => fileName.startsWith('word/media/') && !zip.files[fileName].dir
  );

  for (const filePath of mediaFiles) {
    const file = zip.files[filePath];
    const base64 = await file.async('base64');
    const name = filePath.split('/').pop() || 'image.png';
    images.push({ name, base64 });
  }

  return images;
}
