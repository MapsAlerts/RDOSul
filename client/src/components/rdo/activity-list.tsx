import { useState } from "react";
import { Wand2, Edit2, Check, X, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";

interface Activity {
  id: string;
  text: string;
  originalText: string;
  isRewritten: boolean;
}

interface ActivityListProps {
  activities: Activity[];
  onUpdateActivity: (id: string, newText: string, isRewritten: boolean) => void;
}

export function ActivityList({ activities, onUpdateActivity }: ActivityListProps) {
  return (
    <div className="space-y-4">
      {activities.map((activity) => (
        <ActivityItem 
          key={activity.id} 
          activity={activity} 
          onUpdate={onUpdateActivity} 
        />
      ))}
      
      {activities.length === 0 && (
        <div className="text-center py-12 text-muted-foreground border rounded-lg bg-muted/20">
          Nenhuma atividade encontrada para esta data. Verifique sua planilha ou a data selecionada.
        </div>
      )}
    </div>
  );
}

function ActivityItem({ activity, onUpdate }: { activity: Activity, onUpdate: (id: string, t: string, r: boolean) => void }) {
  const [isEditing, setIsEditing] = useState(false);
  const [editText, setEditText] = useState(activity.text);
  const [isRewriting, setIsRewriting] = useState(false);

  const handleSave = () => {
    onUpdate(activity.id, editText, activity.isRewritten); // Keep rewritten status if just manual edit
    setIsEditing(false);
  };

  const handleRewrite = async () => {
    setIsRewriting(true);
    // Simulate AI delay
    setTimeout(() => {
      // Mock AI improvement: Capitalize, add periods, make it sound "technical"
      const improved = `Execução de ${editText.toLowerCase().replace(/^execução de /, "")}. Atividade realizada conforme procedimentos técnicos e normas de segurança vigentes.`;
      setEditText(improved);
      onUpdate(activity.id, improved, true);
      setIsRewriting(false);
      setIsEditing(false); // Auto-save-ish
    }, 1500);
  };

  return (
    <Card className="overflow-hidden border-l-4 border-l-primary">
      <CardContent className="p-4">
        <div className="flex justify-between items-start gap-4">
          <div className="flex-1">
            <div className="flex items-center gap-2 mb-2">
              <Badge variant="outline" className="text-xs font-mono">ATIVIDADE</Badge>
              {activity.isRewritten && (
                <Badge variant="secondary" className="text-xs gap-1">
                  <Wand2 className="w-3 h-3" /> Melhorado por IA
                </Badge>
              )}
            </div>
            
            {isEditing ? (
              <Textarea 
                value={editText} 
                onChange={(e) => setEditText(e.target.value)} 
                className="min-h-[100px] font-sans text-base"
              />
            ) : (
              <p className="text-sm text-foreground/90 leading-relaxed">
                {activity.text}
              </p>
            )}
          </div>
          
          <div className="flex flex-col gap-2">
            {isEditing ? (
              <>
                <Button size="sm" onClick={handleSave} className="w-full">
                  <Check className="w-4 h-4 mr-2" /> Salvar
                </Button>
                <Button size="sm" variant="ghost" onClick={() => { setIsEditing(false); setEditText(activity.text); }}>
                  <X className="w-4 h-4" />
                </Button>
              </>
            ) : (
              <>
                <Button size="sm" variant="outline" onClick={() => setIsEditing(true)}>
                  <Edit2 className="w-4 h-4 mr-2" /> Editar
                </Button>
                <Button 
                  size="sm" 
                  variant="secondary" 
                  onClick={handleRewrite}
                  disabled={isRewriting}
                >
                  {isRewriting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Wand2 className="w-4 h-4" />}
                </Button>
              </>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
