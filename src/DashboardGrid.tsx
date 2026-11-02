import React from 'react';
import {DndContext, DragOverlay, KeyboardSensor, PointerSensor, closestCenter, useSensor, useSensors, type DragStartEvent, type DragOverEvent, type CollisionDetection} from '@dnd-kit/core';
import {SortableContext, useSortable, arrayMove, rectSortingStrategy, sortableKeyboardCoordinates, defaultAnimateLayoutChanges} from '@dnd-kit/sortable';
import {CSS} from '@dnd-kit/utilities';
import {Tooltip,Button,Modal,TextInput} from '@gravity-ui/uikit';
import {AppIcon} from './Icons';
import {MAX_CARD_TITLE_LENGTH} from './useCardTitles';
import {dashboardGroups} from './dashboardGroups';
import {cardHelp} from './cardHelp';
export type DashboardCard = {id: string; title: string; kind: 'metric' | 'panel'; content: React.ReactNode};
const STORAGE_KEY = 'netting.dashboard.card-order.v2.grouped';
function CardContent({card,onRename}: {card: DashboardCard;onRename?:()=>void}) {
  const children = React.Children.toArray((card.content as React.ReactElement<{children:React.ReactNode}>).props.children);
  const heading = <><span>{card.title}</span><Tooltip content={cardHelp[card.id]} openDelay={150} placement="bottom" className="card-help-tooltip"><button type="button" className="card-help" aria-label={`How it is calculated: ${card.title}`}>?</button></Tooltip>{onRename && <button type="button" className="card-rename" aria-label={`Rename: ${card.title}`} title="Edit card name" onClick={onRename}><AppIcon name="pencil" size={15}/></button>}</>;
  return <>{card.kind==='panel' ? <h2 className="card-heading" aria-label={card.title}>{heading}</h2> : <div className="card-heading">{heading}</div>}{children.slice(1)}</>;
}
export function normalizeOrder(saved: unknown, ids: string[]) {
  const existing = Array.isArray(saved) ? saved.filter((id): id is string => typeof id === 'string' && ids.includes(id)) : [];
  return [...new Set([...existing, ...ids])];
}
function SortableCard({card, loading,onRename}: {card: DashboardCard; loading: boolean;onRename:()=>void}) {
  const {attributes,listeners,setNodeRef,transform,transition,isDragging} = useSortable({id:card.id, disabled:loading, animateLayoutChanges: args=>defaultAnimateLayoutChanges({...args, wasDragging:true})});
  return <article ref={setNodeRef} data-card-id={card.id} aria-busy={loading} className={`dashboard-card ${card.kind} ${card.id==='total-combined-today'?'primary-metric':''} ${isDragging ? 'is-dragging' : ''} ${loading ? 'is-loading' : ''}`} style={{transform:CSS.Transform.toString(transform),transition}}>
    <button disabled={loading} className="drag-handle" {...attributes} {...listeners} aria-label={`Move: ${card.title}`} title="Drag this card. Keyboard: Space, arrow keys, Space."><AppIcon name="grip" size={22}/></button>
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
  const groupForId=(id:string)=>dashboardGroups.find(group=>group.cards.includes(id))?.id ?? 'details';
  const groupedIds=dashboardGroups.flatMap(group=>group.cards).filter(id=>cards.some(card=>card.id===id));
  const ids = [...groupedIds,...cards.filter(card=>!groupedIds.includes(card.id)).map(card=>card.id)];
  const [order,setOrder] = React.useState<string[]>(()=>{try{return normalizeOrder(JSON.parse(localStorage.getItem(storageKey) ?? 'null'),ids);}catch{return ids;}});
  const [active,setActive] = React.useState<string | null>(null);
  const [width,setWidth] = React.useState(0);
  const initial = React.useRef(order);
  const orderRef = React.useRef(order);
  const sensors = useSensors(useSensor(PointerSensor,{activationConstraint:{distance:6}}),useSensor(KeyboardSensor,{coordinateGetter:sortableKeyboardCoordinates}));
  const normalizedOrder = normalizeOrder(order,ids);
  const visibleGroups=dashboardGroups.map(group=>({...group,cards:normalizedOrder.filter(id=>group.cards.includes(id))})).filter(group=>group.cards.length>0);
  const panelIds=normalizedOrder.filter(id=>!groupedIds.includes(id));
  const visibleOrder=[...visibleGroups.flatMap(group=>group.cards),...panelIds];
  const collisions:CollisionDetection=args=>closestCenter({...args,droppableContainers:args.droppableContainers.filter(container=>groupForId(String(container.id))===groupForId(String(args.active.id)))});
  const lookup = new Map(cards.map(card=>[card.id,{...card,title:titles[card.id]??card.title}]));
  const activeCard = active ? lookup.get(active) : undefined;
  function start(event: DragStartEvent) {initial.current=visibleOrder;orderRef.current=visibleOrder;setActive(String(event.active.id));setWidth(event.active.rect.current.initial?.width ?? 300);}
  function over(event: DragOverEvent) {
    if (!event.over || event.active.id === event.over.id || groupForId(String(event.active.id))!==groupForId(String(event.over.id))) return;
    const next=normalizeOrder(orderRef.current,ids);
    const from=next.indexOf(String(event.active.id));const to=next.indexOf(String(event.over.id));
    if(from<0 || to<0)return;
    orderRef.current=arrayMove(next,from,to);setOrder(orderRef.current);
  }
  function end() {const next=normalizeOrder(orderRef.current,ids);setOrder(next);setActive(null);try{localStorage.setItem(storageKey,JSON.stringify(next));}catch{/* Layout remains usable when storage is unavailable. */}}
  function cancel() {orderRef.current=initial.current;setOrder(initial.current);setActive(null);}

  function edit(card:DashboardCard){setDraft(titles[card.id]??card.title);setValidation('');setEditingId(card.id);}
  function close(){setEditingId(null);setValidation('');}
  function save(event:React.FormEvent){event.preventDefault();const title=draft.trim();if(!title){setValidation('Enter a card name.');return;}if(title.length>MAX_CARD_TITLE_LENGTH){setValidation(`Use no more than ${MAX_CARD_TITLE_LENGTH} characters.`);return;}if(editingCard){onRename(editingCard.id,title===editingCard.title?null:title);close();}}
  return <><Modal open={Boolean(editingCard)} onOpenChange={open=>{if(!open)close();}} aria-labelledby="rename-card-title" contentClassName="rename-dialog" initialFocus={titleInput}><form onSubmit={save}><h2 id="rename-card-title">Card name</h2><TextInput size="l" label="Name" value={draft} controlRef={titleInput} controlProps={{'aria-label':'Card name',maxLength:MAX_CARD_TITLE_LENGTH}} onUpdate={value=>{setDraft(value);setValidation('');}} validationState={validation?'invalid':undefined} errorMessage={validation}/><p className="rename-caption">This name is saved on this device and used in every view.</p><div className="rename-reset"><Button type="button" view="flat" onClick={()=>{if(editingCard){onRename(editingCard.id,null);close();}}}>Restore default name</Button></div><div className="dialog-actions"><Button type="button" size="l" onClick={close}>Cancel</Button><Button size="l" view="action" type="submit">Save</Button></div></form></Modal><div className="layout-hint">Drag cards within each block. Your layout is saved automatically.</div><DndContext sensors={sensors} collisionDetection={collisions} onDragStart={start} onDragOver={over} onDragEnd={end} onDragCancel={cancel}>
    <div className="metric-groups">{visibleGroups.map((group,index)=><section className={`metric-group metric-group--${group.id}`} data-group-id={group.id} aria-labelledby={`group-${group.id}`} key={group.id}>
      <div className="metric-group-heading"><span className="group-index" aria-hidden="true">{String(index+1).padStart(2,'0')}</span><h2 id={`group-${group.id}`}>{group.title}</h2></div>
      <SortableContext items={group.cards} strategy={rectSortingStrategy}><div className={`dashboard-grid group-grid group-grid--${group.cards.length}`} role="group" aria-label={`${group.title} cards`}>{group.cards.map(id=><SortableCard key={id} card={lookup.get(id)!} loading={loading} onRename={()=>edit(lookup.get(id)!)}/>)}</div></SortableContext>
    </section>)}</div>
    {panelIds.length>0 && <SortableContext items={panelIds} strategy={rectSortingStrategy}><section className="dashboard-grid dashboard-details" aria-label="Charts and details">{panelIds.map(id=><SortableCard key={id} card={lookup.get(id)!} loading={loading} onRename={()=>edit(lookup.get(id)!)}/>)}</section></SortableContext>}
    <DragOverlay>{activeCard ? <article className={`dashboard-card ${activeCard.kind} drag-preview`} style={{width}}><CardContent card={activeCard}/></article> : null}</DragOverlay>
  </DndContext></>;
}
