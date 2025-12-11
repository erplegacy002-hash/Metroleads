import React from 'react';
import { Download } from 'lucide-react';
import { GeneratedImage } from '../types';

interface ImageGalleryProps {
  images: GeneratedImage[];
}

const ImageGallery: React.FC<ImageGalleryProps> = ({ images }) => {
  if (images.length === 0) return null;

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 w-full max-w-6xl mx-auto animate-fade-in">
      {images.map((img, index) => (
        <div key={index} className="bg-white rounded-xl shadow-md overflow-hidden border border-slate-100 flex flex-col">
          <div className="p-4 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
            <h3 className="font-semibold text-slate-700 truncate" title={img.project_name}>
              {img.project_name}
            </h3>
            <a
              href={img.image_url}
              download={img.filename}
              className="flex items-center space-x-1 text-xs font-medium text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-3 py-1.5 rounded-full transition-colors"
            >
              <Download className="w-3 h-3" />
              <span>Save</span>
            </a>
          </div>
          <div className="p-4 flex-grow flex items-center justify-center bg-slate-200/50 min-h-[200px]">
             {/* Using a simplified path for demo; backend will serve this */}
            <img 
              src={img.image_url} 
              alt={`Report for ${img.project_name}`}
              className="max-w-full h-auto shadow-sm rounded border border-slate-200"
            />
          </div>
        </div>
      ))}
    </div>
  );
};

export default ImageGallery;