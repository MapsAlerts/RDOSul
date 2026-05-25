import { useState } from "react";
import { Upload, FileSpreadsheet, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { cn } from "@/lib/utils";

interface FileUploadProps {
  onFileSelect: (file: File) => void;
  accept: string;
  label: string;
  icon?: React.ReactNode;
}

export function FileUpload({ onFileSelect, accept, label, icon }: FileUploadProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      setSelectedFile(file);
      onFileSelect(file);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      setSelectedFile(file);
      onFileSelect(file);
    }
  };

  const clearFile = (e: React.MouseEvent) => {
    e.stopPropagation();
    setSelectedFile(null);
  };

  return (
    <Card
      className={cn(
        "border-2 border-dashed transition-colors cursor-pointer relative overflow-hidden group",
        isDragging ? "border-secondary bg-secondary/5" : "border-muted hover:border-primary/50",
        selectedFile ? "border-primary bg-primary/5" : ""
      )}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      onClick={() => document.getElementById(`file-input-${label}`)?.click()}
    >
      <CardContent className="flex flex-col items-center justify-center p-6 text-center h-32">
        <input
          id={`file-input-${label}`}
          type="file"
          accept={accept}
          className="hidden"
          onChange={handleFileInput}
        />
        
        {selectedFile ? (
          <div className="flex items-center gap-3 z-10">
            <div className="p-2 bg-primary/10 rounded-full text-primary">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div className="text-left">
              <p className="text-sm font-medium truncate max-w-[180px]">{selectedFile.name}</p>
              <p className="text-xs text-muted-foreground">{(selectedFile.size / 1024).toFixed(1)} KB</p>
            </div>
            <Button variant="ghost" size="icon" onClick={clearFile} className="h-8 w-8 ml-2 text-muted-foreground hover:text-destructive">
              <X className="w-4 h-4" />
            </Button>
          </div>
        ) : (
          <>
            <div className={cn("p-3 rounded-full mb-3 transition-colors", isDragging ? "bg-secondary text-secondary-foreground" : "bg-muted text-muted-foreground group-hover:bg-primary/10 group-hover:text-primary")}>
              {icon || <Upload className="w-5 h-5" />}
            </div>
            <p className="text-sm font-medium text-foreground">{label}</p>
            <p className="text-xs text-muted-foreground mt-1">Clique ou arraste o arquivo aqui</p>
          </>
        )}
      </CardContent>
    </Card>
  );
}
