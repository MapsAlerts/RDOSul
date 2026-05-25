import { HardHat } from "lucide-react";

export function Header() {
  return (
    <header className="border-b bg-card">
      <div className="container mx-auto px-4 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 bg-primary rounded-md flex items-center justify-center text-primary-foreground">
            <HardHat className="w-6 h-6" />
          </div>
          <div>
            <h1 className="font-display font-bold text-xl leading-none text-primary">Gerador de RDO</h1>
          </div>
        </div>
        
        <div className="flex items-center gap-4 text-sm text-muted-foreground">
          <span>v2.4.0 (Web)</span>
          <div className="h-4 w-px bg-border" />
          <span className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-green-500" />
            Sistema Online
          </span>
        </div>
      </div>
    </header>
  );
}
