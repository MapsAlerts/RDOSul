import { useState } from "react";
import { Upload, X, Image as ImageIcon, Link as LinkIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

export interface ImageItem {
  id: string;
  file?: File;
  url?: string;
  preview: string;
}

interface ImageSectionProps {
  title: string;
  images: ImageItem[];
  onAddImage: (file: File) => void;
  onRemoveImage: (id: string) => void;
}

export function ImageSection({ title, images, onAddImage, onRemoveImage }: ImageSectionProps) {
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      Array.from(e.target.files).forEach(file => onAddImage(file));
    }
  };

  return (
    <Card className="h-full flex flex-col">
      <CardHeader className="pb-3 bg-muted/30 border-b">
        <CardTitle className="text-sm font-medium uppercase tracking-wider text-muted-foreground flex justify-between items-center">
          {title}
          <span className="text-xs bg-primary/10 text-primary px-2 py-0.5 rounded-full">{images.length}</span>
        </CardTitle>
      </CardHeader>
      <CardContent className="flex-1 p-4 space-y-4">
        <div className="grid grid-cols-2 gap-3">
          {images.map((img) => (
            <div key={img.id} className="relative group aspect-square rounded-md overflow-hidden border bg-muted">
              <img src={img.preview} alt="Evidência" className="w-full h-full object-cover" />
              
              {/* Indicator for URL vs File */}
              {img.url && (
                 <div className="absolute top-1 right-1 bg-black/50 p-1 rounded-full text-white">
                   <LinkIcon className="w-3 h-3" />
                 </div>
              )}

              <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                <Button 
                  variant="destructive" 
                  size="icon" 
                  className="h-8 w-8 rounded-full"
                  onClick={() => onRemoveImage(img.id)}
                >
                  <X className="w-4 h-4" />
                </Button>
              </div>
            </div>
          ))}
          
          <label className="aspect-square rounded-md border-2 border-dashed border-muted hover:border-primary/50 hover:bg-muted/50 transition-all flex flex-col items-center justify-center cursor-pointer text-muted-foreground hover:text-primary">
            <input type="file" accept="image/*" multiple className="hidden" onChange={handleFileChange} />
            <Upload className="w-6 h-6 mb-2" />
            <span className="text-xs font-medium">Adicionar Foto</span>
          </label>
        </div>
      </CardContent>
    </Card>
  );
}
