import React from 'react';
import {DndContext, DragOverlay, KeyboardSensor, PointerSensor, closestCenter, useSensor, useSensors, type DragStartEvent, type DragOverEvent} from '@dnd-kit/core';
import {SortableContext, useSortable, arrayMove, rectSortingStrategy, sortableKeyboardCoordinates, defaultAnimateLayoutChanges} from '@dnd-kit/sortable';
import {CSS} from '@dnd-kit/utilities';
import {Tooltip,Button,Modal,TextInput} from '@gravity-ui/uikit';
import {AppIcon} from './Icons';
import {MAX_CARD_TITLE_LENGTH} from './useCardTitles';
import {cardHelp} from './cardHelp';
export type DashboardCard = {id: string; title: string; kind: 'metric' | 'panel'; content: React.ReactNode};
const STORAGE_KEY = 'netting.dashboard.card-order.v1';
function CardContent({card,onRename}: {card: DashboardCard;onRename?:()=>void}) {
  const children = React.Children.toArray((card.content as React.ReactElement<{children:React.ReactNode}>).props.children);
  const heading = <><span>{card.title}</span><Tooltip content={cardHelp[card.id]} openDelay={150} placement="bottom" className="card-help-tooltip"><button type="button" className="card-help" aria-label={`Как рассчитывается: ${card.title}`}>?</button></Tooltip>{onRename && <button type="button" className="card-rename" aria-label={`Переименовать: ${card.title}`} title="Изменить название карточки" onClick={onRename}><AppIcon name="pencil" size={15}/></button>}</>;
  return <>{card.kind==='panel' ? <h2 className="card-heading" aria-label={card.title}>{heading}</h2> : <div className="card-heading">{heading}</div>}{children.slice(1)}</>;
}
export function normalizeOrder(saved: unknown, ids: string[]) {
  const existing = Array.isArray(saved) ? saved.filter((id): id is string => typeof id === 'string' && ids.includes(id)) : [];
  return [...new Set([...existing, ...ids])];
}
function SortableCard({card, loading,onRename}: {card: DashboardCard; loading: boolean;onRename:()=>void}) {
  const {attributes,listeners,setNodeRef,transform,transition,isDragging} = useSortable({id:card.id, disabled:loading, animateLayoutChanges: args=>defaultAnimateLayoutChanges({...args, wasDragging:true})});
  return <article ref={setNodeRef} aria-busy={loading} className={`dashboard-card ${card.kind} ${isDragging ? 'is-dragging' : ''} ${loading ? 'is-loading' : ''}`} style={{transform:CSS.Transform.toString(transform),transition}}>
    <button disabled={loading} className="drag-handle" {...attributes} {...listeners} aria-label={`Переместить: ${card.title}`} title="Перетащите карточку. С клавиатуры: пробел, стрелки, пробел."><AppIcon name="grip" size={22}/></button>
    <div className="card-content" aria-hidden={loading || undefined}><CardContent card={card} onRename={loading?undefined:onRename}/></div>
    {loading && <div className={`card-skeleton ${card.kind}`} aria-hidden="true"><span className="skeleton-bar skeleton-title"/>{card.kind === 'metric' ? <><span className="skeleton-bar skeleton-value"/><span className="skeleton-bar skeleton-caption"/></> : <><span className="skeleton-bar skeleton-caption"/><div className="skeleton-panel"/></>}</div>}
  </article>;
}
export function DashboardGrid({cards, loading = false, layoutId = 'overview',titles,onRename}: {cards: DashboardCard[]; loading?: boolean; layoutId?: string;titles:Record<string,string>;onRename:(id:string,title:string|null)=>void}) {
  const storageKey = layoutId === 'overview' ? STORAGE_KEY : STORAGE_KEY + '.' + layoutId;
  const [editingId,setEditingId]=React.useState<string|null>(null);
  const [draft,setDraft]=React.useState('');
  const [validation,setValidation]=React.useState('');
  const titleInput=React.useRef<HTMLInputElement>(null);
  const editingCard=cards.find(card=>card.id===editingId);
  const ids = cards.map(c=>c.id);
  const [order,setOrder] = React.useState<string[]>(()=>{try{return normalizeOrder(JSON.parse(localStorage.getItem(storageKey) ?? 'null'),ids);}catch{return ids;}});
  const [active,setActive] = React.useState<string | null>(null);
  const [width,setWidth] = React.useState(0);
  const initial = React.useRef(order);
  const sensors = useSensors(useSensor(PointerSensor,{activationConstraint:{distance:6}}),useSensor(KeyboardSensor,{coordinateGetter:sortableKeyboardCoordinates}));
  const visibleOrder = normalizeOrder(order,ids);
  const lookup = new Map(cards.map(card=>[card.id,{...card,title:titles[card.id]??card.title}]));
  const activeCard = active ? lookup.get(active) : undefined;
  function start(event: DragStartEvent) {initial.current=visibleOrder;setActive(String(event.active.id));setWidth(event.active.rect.current.initial?.width ?? 300);}
  function over(event: DragOverEvent) {
    if (!event.over || event.active.id === event.over.id) return;
    setOrder(current=>{const next=normalizeOrder(current,ids);const from=next.indexOf(String(event.active.id));const to=next.indexOf(String(event.over!.id));return from<0 || to<0 ? next : arrayMove(next,from,to);});
  }
  function end() {setActive(null);try{localStorage.setItem(storageKey,JSON.stringify(normalizeOrder(order,ids)));}catch{/* Layout remains usable when storage is unavailable. */}}
  function edit(card:DashboardCard){setDraft(titles[card.id]??card.title);setValidation('');setEditingId(card.id);}
  function close(){setEditingId(null);setValidation('');}
  function save(event:React.FormEvent){event.preventDefault();const title=draft.trim();if(!title){setValidation('Введите название карточки.');return;}if(title.length>MAX_CARD_TITLE_LENGTH){setValidation(`Не больше ${MAX_CARD_TITLE_LENGTH} символов.`);return;}if(editingCard){onRename(editingCard.id,title===editingCard.title?null:title);close();}}
  return <><Modal open={Boolean(editingCard)} onOpenChange={open=>{if(!open)close();}} aria-labelledby="rename-card-title" contentClassName="rename-dialog" initialFocus={titleInput}><form onSubmit={save}><h2 id="rename-card-title">Название карточки</h2><TextInput size="l" label="Название" value={draft} controlRef={titleInput} controlProps={{'aria-label':'Название карточки',maxLength:MAX_CARD_TITLE_LENGTH}} onUpdate={value=>{setDraft(value);setValidation('');}} validationState={validation?'invalid':undefined} errorMessage={validation}/><p className="rename-caption">Название сохраняется на этом устройстве и применяется во всех разделах.</p><div className="rename-reset"><Button type="button" view="flat" onClick={()=>{if(editingCard){onRename(editingCard.id,null);close();}}}>Вернуть исходное название</Button></div><div className="dialog-actions"><Button type="button" size="l" onClick={close}>Отмена</Button><Button size="l" view="action" type="submit">Сохранить</Button></div></form></Modal><div className="layout-hint">Перетаскивайте карточки за ⠿ — порядок сохраняется автоматически.</div><DndContext sensors={sensors} collisionDetection={closestCenter} onDragStart={start} onDragOver={over} onDragEnd={end} onDragCancel={()=>{setOrder(initial.current);setActive(null);}}>
    <SortableContext items={visibleOrder} strategy={rectSortingStrategy}><section className="dashboard-grid" aria-label="Карточки дашборда">{visibleOrder.map(id=><SortableCard key={id} card={lookup.get(id)!} loading={loading} onRename={()=>edit(lookup.get(id)!)}/>)}</section></SortableContext>
    <DragOverlay>{activeCard ? <article className={`dashboard-card ${activeCard.kind} drag-preview`} style={{width}}><CardContent card={activeCard}/></article> : null}</DragOverlay>
  </DndContext></>;
}
