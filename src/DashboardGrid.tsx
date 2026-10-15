import React from 'react';
import {DndContext, DragOverlay, KeyboardSensor, PointerSensor, closestCenter, useSensor, useSensors, type DragStartEvent, type DragOverEvent} from '@dnd-kit/core';
import {SortableContext, useSortable, arrayMove, rectSortingStrategy, sortableKeyboardCoordinates, defaultAnimateLayoutChanges} from '@dnd-kit/sortable';
import {CSS} from '@dnd-kit/utilities';
import {Tooltip} from '@gravity-ui/uikit';
import {AppIcon} from './Icons';
import {cardHelp} from './cardHelp';
export type DashboardCard = {id: string; title: string; kind: 'metric' | 'panel'; content: React.ReactNode};
const STORAGE_KEY = 'netting.dashboard.card-order.v1';
function CardContent({card}: {card: DashboardCard}) {
  const children = React.Children.toArray((card.content as React.ReactElement<{children:React.ReactNode}>).props.children);
  const heading = <><span>{card.title}</span><Tooltip content={cardHelp[card.id]} openDelay={150} placement="bottom" className="card-help-tooltip"><button type="button" className="card-help" aria-label={`Как рассчитывается: ${card.title}`}>?</button></Tooltip></>;
  return <>{card.kind==='panel' ? <h2 className="card-heading" aria-label={card.title}>{heading}</h2> : <div className="card-heading">{heading}</div>}{children.slice(1)}</>;
}
export function normalizeOrder(saved: unknown, ids: string[]) {
  const existing = Array.isArray(saved) ? saved.filter((id): id is string => typeof id === 'string' && ids.includes(id)) : [];
  return [...new Set([...existing, ...ids])];
}
function SortableCard({card, loading}: {card: DashboardCard; loading: boolean}) {
  const {attributes,listeners,setNodeRef,transform,transition,isDragging} = useSortable({id:card.id, disabled:loading, animateLayoutChanges: args=>defaultAnimateLayoutChanges({...args, wasDragging:true})});
  return <article ref={setNodeRef} aria-busy={loading} className={`dashboard-card ${card.kind} ${isDragging ? 'is-dragging' : ''} ${loading ? 'is-loading' : ''}`} style={{transform:CSS.Transform.toString(transform),transition}}>
    <button disabled={loading} className="drag-handle" {...attributes} {...listeners} aria-label={`Переместить: ${card.title}`} title="Перетащите карточку. С клавиатуры: пробел, стрелки, пробел."><AppIcon name="grip" size={22}/></button>
    <div className="card-content" aria-hidden={loading || undefined}><CardContent card={card}/></div>
    {loading && <div className={`card-skeleton ${card.kind}`} aria-hidden="true"><span className="skeleton-bar skeleton-title"/>{card.kind === 'metric' ? <><span className="skeleton-bar skeleton-value"/><span className="skeleton-bar skeleton-caption"/></> : <><span className="skeleton-bar skeleton-caption"/><div className="skeleton-panel"/></>}</div>}
  </article>;
}
export function DashboardGrid({cards, loading = false, layoutId = 'overview'}: {cards: DashboardCard[]; loading?: boolean; layoutId?: string}) {
  const storageKey = layoutId === 'overview' ? STORAGE_KEY : STORAGE_KEY + '.' + layoutId;
  const ids = cards.map(c=>c.id);
  const [order,setOrder] = React.useState<string[]>(()=>{try{return normalizeOrder(JSON.parse(localStorage.getItem(storageKey) ?? 'null'),ids);}catch{return ids;}});
  const [active,setActive] = React.useState<string | null>(null);
  const [width,setWidth] = React.useState(0);
  const initial = React.useRef(order);
  const sensors = useSensors(useSensor(PointerSensor,{activationConstraint:{distance:6}}),useSensor(KeyboardSensor,{coordinateGetter:sortableKeyboardCoordinates}));
  const visibleOrder = normalizeOrder(order,ids);
  const lookup = new Map(cards.map(c=>[c.id,c]));
  const activeCard = active ? lookup.get(active) : undefined;
  function start(event: DragStartEvent) {initial.current=visibleOrder;setActive(String(event.active.id));setWidth(event.active.rect.current.initial?.width ?? 300);}
  function over(event: DragOverEvent) {
    if (!event.over || event.active.id === event.over.id) return;
    setOrder(current=>{const next=normalizeOrder(current,ids);const from=next.indexOf(String(event.active.id));const to=next.indexOf(String(event.over!.id));return from<0 || to<0 ? next : arrayMove(next,from,to);});
  }
  function end() {setActive(null);try{localStorage.setItem(storageKey,JSON.stringify(normalizeOrder(order,ids)));}catch{/* Layout remains usable when storage is unavailable. */}}
  return <><div className="layout-hint">Перетаскивайте карточки за ⠿ — порядок сохраняется автоматически.</div><DndContext sensors={sensors} collisionDetection={closestCenter} onDragStart={start} onDragOver={over} onDragEnd={end} onDragCancel={()=>{setOrder(initial.current);setActive(null);}}>
    <SortableContext items={visibleOrder} strategy={rectSortingStrategy}><section className="dashboard-grid" aria-label="Карточки дашборда">{visibleOrder.map(id=><SortableCard key={id} card={lookup.get(id)!} loading={loading}/>)}</section></SortableContext>
    <DragOverlay>{activeCard ? <article className={`dashboard-card ${activeCard.kind} drag-preview`} style={{width}}><CardContent card={activeCard}/></article> : null}</DragOverlay>
  </DndContext></>;
}
