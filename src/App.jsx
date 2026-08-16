import { Children, useEffect, useMemo, useRef, useState } from "react";
import {
  CalendarDots, CaretDown, CaretLeft, CaretRight, CaretUp, Check, CheckCircle, CircleNotch, DownloadSimple,
  BookOpen, BookmarkSimple, CalendarPlus, DotsSixVertical, FileImage, FileXls, GearSix, CoatHanger as Hanger, Heart, ImageSquare, Link as LinkIcon,
  ArrowsDownUp, ListBullets, MagnifyingGlass, Moon, NotePencil, PencilSimple, Plus, SquaresFour, StackSimple, Sun, TShirt, Trash, UploadSimple, X,
} from "@phosphor-icons/react";
import {
  buildMonth, createBackupPayload, createWardrobeResetState, filterItemsByCategory, formatFileTimestamp, getDetailWidthLimits, iso, moveItemByOffset,
  nextLookbookName, parseWardrobeExcelRows, removeItemReferences, reorderItemIds, searchLookbooks, searchWardrobeItems, searchWishlistItems, validItemIds, validateBackupPayload, wardrobeItemsToExcelRows,
  WARDROBE_EXCEL_HEADERS,
} from "./domain.mjs";

const TODAY = iso(new Date());
const DEFAULT_CATEGORIES = ["상의", "하의", "아우터", "원피스", "신발", "가방", "액세서리"];
const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
const YEAR_OPTIONS = Array.from({ length: 21 }, (_, index) => 2020 + index);
const MONTH_OPTIONS = Array.from({ length: 12 }, (_, index) => index + 1);
const TEMPLATE_URL = "/templates/MyCloset_옷장_업로드_템플릿.xlsx";
const DESKTOP_UI_DENSITY = 0.9;
const PRODUCT_LOOKUP_TIMEOUT_MS = 15_000;
const PRODUCT_LOOKUP_TIMEOUT_MESSAGE = "상품 정보를 가져오는 데 시간이 오래 걸리고 있어요. 직접 입력으로 등록해주세요.";
const INITIAL_ITEMS = [];
const INITIAL_OUTFITS = {};
const INITIAL_WISHLIST = [];
const INITIAL_LOOKBOOKS = [];
const INITIAL_INSPIRATIONS = [];
const DEMO_ITEM_IDS = new Set(["shirt-blue", "pants-black", "bag-black", "shoes-black", "shirt-beige", "pants-photo", "cardigan-navy", "tee-brown", "denim"]);
const DEMO_LOOKBOOK_IDS = new Set(["lookbook-daily", "lookbook-weekend"]);
const DEMO_WISHLIST_IDS = new Set(["wish-1", "wish-2"]);
const DEMO_OUTFIT_DATES = new Set(["2026-08-03", "2026-08-04", "2026-08-06", "2026-08-08", "2026-08-10", "2026-08-12", "2026-08-14", "2026-08-18", "2026-08-20", "2026-08-22", "2026-08-26", "2026-08-28", "2026-08-31"]);
const DATA_CLEANUP_VERSION = 1;

function load(key, fallback) {
  try { const value = localStorage.getItem(key); return value ? JSON.parse(value) : fallback; }
  catch { return fallback; }
}
function migrateDemoData() {
  if (typeof localStorage === "undefined" || Number(load("mycloset-data-cleanup-version", 0)) >= DATA_CLEANUP_VERSION) return;
  const nextItems = load("mycloset-items", []).filter((item) => !DEMO_ITEM_IDS.has(item.id));
  const validIds = new Set(nextItems.map((item) => item.id));
  const nextOutfits = Object.fromEntries(Object.entries(load("mycloset-outfits", {}))
    .map(([date, ids]) => [date, Array.isArray(ids) ? ids.filter((id) => !DEMO_ITEM_IDS.has(id) && validIds.has(id)) : []])
    .filter(([, ids]) => ids.length));
  const nextNotes = Object.fromEntries(Object.entries(load("mycloset-outfit-notes", {})).filter(([date]) => !DEMO_OUTFIT_DATES.has(date)));
  const nextLookbooks = load("mycloset-lookbooks", []).filter((entry) => !DEMO_LOOKBOOK_IDS.has(entry.id)).map((entry) => ({ ...entry, itemIds: entry.itemIds.filter((id) => validIds.has(id)) })).filter((entry) => entry.itemIds.length);
  const nextWishlist = load("mycloset-wishlist", []).filter((entry) => !DEMO_WISHLIST_IDS.has(entry.id));
  localStorage.setItem("mycloset-items", JSON.stringify(nextItems));
  localStorage.setItem("mycloset-outfits", JSON.stringify(nextOutfits));
  localStorage.setItem("mycloset-outfit-notes", JSON.stringify(nextNotes));
  localStorage.setItem("mycloset-lookbooks", JSON.stringify(nextLookbooks));
  localStorage.setItem("mycloset-wishlist", JSON.stringify(nextWishlist));
  localStorage.setItem("mycloset-data-cleanup-version", JSON.stringify(DATA_CLEANUP_VERSION));
}
migrateDemoData();
function detailWidthLimits(sidebarCollapsed) {
  if (typeof window === "undefined") return { minWidth: 360, maxWidth: 1120 };
  return getDetailWidthLimits(window.innerWidth, sidebarCollapsed);
}
async function fileToDataUrl(file) {
  if (!file?.size) return "";
  if (file.size > 2.5 * 1024 * 1024) throw new Error("이미지는 2.5MB 이하만 등록할 수 있어요.");
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("이미지를 읽지 못했어요."));
    reader.readAsDataURL(file);
  });
}
async function optimizeScrapImage(file) {
  if (!file?.type?.startsWith("image/")) throw new Error("이미지 파일만 붙여넣거나 선택할 수 있어요.");
  if (file.size > 2.5 * 1024 * 1024) throw new Error("이미지는 2.5MB 이하만 등록할 수 있어요.");
  const objectUrl = URL.createObjectURL(file);
  try {
    const image = await new Promise((resolve, reject) => {
      const preview = new Image();
      preview.onload = () => resolve(preview);
      preview.onerror = () => reject(new Error("이미지를 읽지 못했어요."));
      preview.src = objectUrl;
    });
    const maxDimension = 1600;
    let scale = Math.min(1, maxDimension / Math.max(image.naturalWidth, image.naturalHeight));
    let result = "";
    for (let attempt = 0; attempt < 4; attempt += 1) {
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
      const context = canvas.getContext("2d");
      if (!context) throw new Error("이미지를 최적화하지 못했어요.");
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      result = canvas.toDataURL("image/webp", Math.max(.58, .82 - attempt * .08));
      if (result.length <= 850_000) break;
      scale *= .82;
    }
    return result;
  } finally { URL.revokeObjectURL(objectUrl); }
}
function writeLocalStorageSnapshot(entries) {
  const serializedEntries = entries.map(([key, value]) => [key, JSON.stringify(value)]);
  const previousValues = new Map(serializedEntries.map(([key]) => [key, localStorage.getItem(key)]));
  try {
    serializedEntries.forEach(([key, value]) => localStorage.setItem(key, value));
  } catch (error) {
    serializedEntries.forEach(([key]) => localStorage.removeItem(key));
    previousValues.forEach((value, key) => { if (value !== null) localStorage.setItem(key, value); });
    throw error;
  }
}
function Required() { return <span className="required-mark" aria-label="필수">*</span>; }
function FieldTitle({ children, required = false }) { return <span className="field-title">{children}{required && <Required />}</span>; }
function CustomSelect({ name, value, defaultValue, onValueChange, children, required = false, label = "옵션 선택", variant = "field", className = "" }) {
  const options = Children.toArray(children).map((child) => ({
    value: String(child.props.value ?? child.props.children),
    label: child.props.children,
    text: Children.toArray(child.props.children).join(""),
    disabled: Boolean(child.props.disabled),
  }));
  const controlled = value !== undefined;
  const [internalValue, setInternalValue] = useState(() => String(defaultValue ?? options[0]?.value ?? ""));
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const triggerRef = useRef(null);
  const optionRefs = useRef([]);
  const selectedValue = String(controlled ? value : internalValue);
  const selectedOption = options.find((option) => option.value === selectedValue) ?? options[0];

  useEffect(() => {
    if (!open) return undefined;
    const closeOnOutsideClick = (event) => {
      if (!rootRef.current?.contains(event.target)) setOpen(false);
    };
    document.addEventListener("pointerdown", closeOnOutsideClick);
    const selectedIndex = Math.max(0, options.findIndex((option) => option.value === selectedValue));
    const frame = requestAnimationFrame(() => optionRefs.current[selectedIndex]?.focus());
    return () => {
      document.removeEventListener("pointerdown", closeOnOutsideClick);
      cancelAnimationFrame(frame);
    };
  }, [open, selectedValue]);

  const choose = (option) => {
    if (option.disabled) return;
    if (!controlled) setInternalValue(option.value);
    onValueChange?.(option.value);
    setOpen(false);
    triggerRef.current?.focus();
  };
  const handleKeyDown = (event) => {
    if (event.key === "Escape" && open) {
      event.preventDefault();
      setOpen(false);
      triggerRef.current?.focus();
      return;
    }
    if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return;
    event.preventDefault();
    if (!open) { setOpen(true); return; }
    const enabledIndexes = options.map((option, index) => option.disabled ? -1 : index).filter((index) => index >= 0);
    const focusedIndex = optionRefs.current.findIndex((option) => option === document.activeElement);
    const currentPosition = Math.max(0, enabledIndexes.indexOf(focusedIndex));
    const nextPosition = event.key === "Home" ? 0 : event.key === "End" ? enabledIndexes.length - 1 : (currentPosition + (event.key === "ArrowDown" ? 1 : -1) + enabledIndexes.length) % enabledIndexes.length;
    optionRefs.current[enabledIndexes[nextPosition]]?.focus();
  };

  const selectedLabel = selectedOption?.label ?? label;
  const selectedText = selectedOption?.text ?? label;
  return <div ref={rootRef} className={`custom-select custom-select-${variant}${open ? " is-open" : ""}${className ? ` ${className}` : ""}`} onKeyDown={handleKeyDown}>
    {name && <input type="hidden" name={name} value={selectedOption?.value ?? ""} required={required}/>}
    <button ref={triggerRef} type="button" className="custom-select-trigger" aria-haspopup="listbox" aria-expanded={open} aria-label={`${label}: ${selectedText}`} title={`${label}: ${selectedText}`} onClick={() => setOpen((current) => !current)}>
      {variant === "sort" ? <ArrowsDownUp size={18}/> : <><span>{selectedLabel}</span><CaretDown size={variant === "compact" ? 14 : 17} weight="bold"/></>}
    </button>
    {open && <div className="custom-select-menu" role="listbox" aria-label={label}>
      {options.map((option, index) => <button ref={(node) => { optionRefs.current[index] = node; }} type="button" className={`custom-select-option${option.value === selectedValue ? " selected" : ""}`} role="option" aria-selected={option.value === selectedValue} disabled={option.disabled} onClick={() => choose(option)} key={option.value}><span>{option.label}</span>{option.value === selectedValue && <Check size={15} weight="bold"/>}</button>)}
    </div>}
  </div>;
}
function SelectField({ name, defaultValue, children, required = false }) {
  return <CustomSelect name={name} defaultValue={defaultValue} required={required} label="분류 선택" className="select-wrap">{children}</CustomSelect>;
}
function SortControl({ label, value, onChange, children }) {
  return <CustomSelect value={value} onValueChange={onChange} label={label} variant="sort" className="sort-control">{children}</CustomSelect>;
}
function SearchField({ value, onChange, label, placeholder }) {
  return <label className="collection-search-field">
    <span className="sr-only">{label}</span>
    <MagnifyingGlass size={18} aria-hidden="true"/>
    <input type="search" value={value} onChange={(event) => onChange(event.target.value)} onKeyDown={(event) => { if (event.key === "Escape" && value) { event.preventDefault(); onChange(""); } }} aria-label={label} placeholder={placeholder}/>
    {value && <button type="button" onClick={() => onChange("")} aria-label={`${label} 지우기`} title="검색어 지우기"><X size={15} weight="bold"/></button>}
  </label>;
}
function SearchEmptyState({ subject, query, category = "전체", onClear, onShowAll }) {
  const hasQuery = Boolean(query.trim());
  return <section className="search-empty-state" role="status">
    <MagnifyingGlass size={28} weight="light"/>
    <div><h2>{hasQuery ? `“${query.trim()}”와 일치하는 ${subject}이(가) 없어요` : `이 카테고리에 ${subject}이(가) 없어요`}</h2><p>검색어나 카테고리를 바꾸면 다른 결과를 확인할 수 있어요.</p></div>
    <div>{hasQuery && <button type="button" className="secondary-button compact" onClick={onClear}>검색어 지우기</button>}{category !== "전체" && <button type="button" className="secondary-button compact" onClick={onShowAll}>전체 카테고리에서 보기</button>}</div>
  </section>;
}
function CompactCalendar({ cursor, onCursorChange, selectedDate, onSelectDate, outfits = {}, className = "", id, role, ariaLabel = "날짜 선택" }) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const days = useMemo(() => buildMonth(cursor), [cursor]);
  const monthTitle = `${cursor.getFullYear()}년 ${cursor.getMonth() + 1}월`;
  const changeMonth = (offset) => {
    setPickerOpen(false);
    onCursorChange(new Date(cursor.getFullYear(), cursor.getMonth() + offset, 1));
  };
  const selectToday = () => {
    const today = new Date();
    setPickerOpen(false);
    onCursorChange(new Date(today.getFullYear(), today.getMonth(), 1));
    onSelectDate(TODAY);
  };
  return <section className={`compact-calendar${className ? ` ${className}` : ""}`} id={id} role={role} aria-label={ariaLabel}>
    <header className="compact-calendar-header"><p aria-live="polite">{monthTitle}</p><div className="month-controls"><div className="month-picker-wrap"><button type="button" className="icon-button" onClick={() => setPickerOpen((current) => !current)} aria-label="연월 선택" aria-expanded={pickerOpen}><CalendarDots size={18}/></button>{pickerOpen && <div className="month-popover" role="dialog" aria-label="연월 선택"><div className="month-popover-head"><strong>연월 선택</strong><CustomSelect value={cursor.getFullYear()} onValueChange={(year) => onCursorChange(new Date(Number(year), cursor.getMonth(), 1))} label="연도 선택" variant="compact" className="month-year-select">{YEAR_OPTIONS.map((year) => <option key={year} value={year}>{year}년</option>)}</CustomSelect></div><div className="month-option-grid">{MONTH_OPTIONS.map((month) => <button type="button" className={cursor.getMonth() + 1 === month ? "active" : ""} aria-pressed={cursor.getMonth() + 1 === month} onClick={() => { onCursorChange(new Date(cursor.getFullYear(), month - 1, 1)); setPickerOpen(false); }} key={month}>{month}월</button>)}</div></div>}</div><button type="button" className="icon-button" onClick={() => changeMonth(-1)} aria-label="이전 달"><CaretLeft size={18}/></button><button type="button" className="today-button" onClick={selectToday} aria-label="오늘 날짜로 이동" title="오늘 날짜로 이동">오늘</button><button type="button" className="icon-button" onClick={() => changeMonth(1)} aria-label="다음 달"><CaretRight size={18}/></button></div></header>
    <div className="compact-calendar-grid"><div className="calendar-grid weekday-row">{WEEKDAYS.map((day, index) => <div className={index === 0 ? "sunday" : ""} key={day}>{day}</div>)}</div><div className="calendar-grid month-grid">{days.map(({ date, key, current }) => { const hasOutfit = (outfits[key] ?? []).length > 0; return <button type="button" className={`day-cell${current ? "" : " outside"}${key === selectedDate ? " selected" : ""}${hasOutfit ? " has-outfit" : ""}`} key={key} onClick={() => onSelectDate(key)} aria-label={`${key}${hasOutfit ? ", 착장 있음" : ", 착장 없음"}`} title={hasOutfit ? "착장 있음" : "착장 없음"}><span className="day-number-row"><span className="day-number">{date.getDate()}</span></span>{hasOutfit && <span className="outfit-indicator" aria-hidden="true"/>}</button>; })}</div></div>
  </section>;
}
const koreanCollator = new Intl.Collator("ko", { numeric: true, sensitivity: "base" });
function sortText(entries, getValue) {
  return [...entries].sort((a, b) => koreanCollator.compare(String(getValue(a) || ""), String(getValue(b) || "")));
}
function sortWardrobe(entries, sort, categoryOrder) {
  if (sort === "oldest") return [...entries];
  if (sort === "name") return sortText(entries, (item) => item.name);
  if (sort === "brand") return [...entries].sort((a, b) => {
    if (!a.brand && b.brand) return 1;
    if (a.brand && !b.brand) return -1;
    return koreanCollator.compare(a.brand || "", b.brand || "") || koreanCollator.compare(a.name || "", b.name || "");
  });
  if (sort === "category") return [...entries].sort((a, b) => {
    const aIndex = categoryOrder.indexOf(a.category);
    const bIndex = categoryOrder.indexOf(b.category);
    return (aIndex < 0 ? categoryOrder.length : aIndex) - (bIndex < 0 ? categoryOrder.length : bIndex) || koreanCollator.compare(a.name || "", b.name || "");
  });
  return [...entries].reverse();
}
function sortLookbookEntries(entries, sort) {
  if (sort === "oldest") return [...entries];
  if (sort === "name") return sortText(entries, (entry) => entry.name);
  if (sort === "items") return [...entries].sort((a, b) => b.itemIds.length - a.itemIds.length || koreanCollator.compare(a.name || "", b.name || ""));
  return [...entries].reverse();
}
function sortWishlistEntries(entries, sort) {
  if (sort === "oldest") return [...entries];
  if (sort === "name") return sortText(entries, (entry) => entry.name);
  if (sort === "brand") return [...entries].sort((a, b) => {
    if (!a.brand && b.brand) return 1;
    if (a.brand && !b.brand) return -1;
    return koreanCollator.compare(a.brand || "", b.brand || "") || koreanCollator.compare(a.name || "", b.name || "");
  });
  return [...entries].reverse();
}
function FilePicker({ id, name, accept, fileName, onChange, label = "파일 선택", icon = "image", compact = false, tone = "secondary" }) {
  return <div className={`custom-file-field${compact ? " settings-file-picker" : ""}`}>
    <input id={id} className="visually-hidden-input" name={name} type="file" accept={accept} onChange={onChange}/>
    <label className={compact ? `${tone}-button compact file-picker-compact-button` : "file-picker-button"} htmlFor={id}>{icon === "excel" ? <FileXls size={18}/> : <UploadSimple size={18}/>} {label}</label>
    <span className={fileName ? "file-name selected" : "file-name"}>{fileName || "선택된 파일 없음"}</span>
  </div>;
}
function ItemVisual({ item }) {
  const [failed, setFailed] = useState(false);
  return item?.image && !failed
    ? <img className="item-image" src={item.image} alt={item.name} draggable="false" onError={() => setFailed(true)}/>
    : <div className="placeholder-visual" role="img" aria-label={`${item?.name ?? "옷"} 이미지 없음`}><TShirt size={34} weight="light"/><span>이미지 없음</span></div>;
}
function Modal({ title, subtitle = "", eyebrow = "MyCloset", onClose, children, wide = false, compact = false, className = "" }) {
  useEffect(() => {
    const onKeyDown = (event) => { if (event.key === "Escape") onClose(); };
    document.addEventListener("keydown", onKeyDown);
    document.body.classList.add("modal-open");
    return () => { document.removeEventListener("keydown", onKeyDown); document.body.classList.remove("modal-open"); };
  }, [onClose]);
  return <div className="modal-backdrop" onMouseDown={onClose}>
    <section className={`modal${wide ? " modal-wide" : ""}${compact ? " modal-compact" : ""}${className ? ` ${className}` : ""}`} role="dialog" aria-modal="true" aria-label={title} onMouseDown={(event) => event.stopPropagation()}>
      <header className="modal-header"><div>{eyebrow && <span className="eyebrow">{eyebrow}</span>}<h2>{title}</h2>{subtitle && <p className="modal-subtitle">{subtitle}</p>}</div><button className="icon-button" onClick={onClose} aria-label="닫기"><X size={20}/></button></header>
      {children}
    </section>
  </div>;
}

function InspirationViewer({ inspiration, onClose }) {
  useEffect(() => {
    const onKeyDown = (event) => { if (event.key === "Escape") onClose(); };
    document.addEventListener("keydown", onKeyDown);
    document.body.classList.add("modal-open");
    return () => { document.removeEventListener("keydown", onKeyDown); document.body.classList.remove("modal-open"); };
  }, [onClose]);
  return <div className="modal-backdrop inspiration-viewer-backdrop" onMouseDown={onClose}>
    <figure className="inspiration-viewer" role="dialog" aria-modal="true" aria-label="스크랩 이미지 크게 보기" onMouseDown={(event) => event.stopPropagation()}>
      <div className="inspiration-viewer-image">
        <img src={inspiration.image} alt={inspiration.caption || "저장한 코디 참고 이미지"}/>
        <button type="button" className="inspiration-viewer-close" onClick={onClose} aria-label="닫기"><X size={22}/></button>
      </div>
      {inspiration.caption && <figcaption>{inspiration.caption}</figcaption>}
    </figure>
  </div>;
}

function ProductUrlStep({ kind, loading, error, onSubmit, onManual }) {
  const subject = kind === "item" ? "옷" : "위시리스트 아이템";
  return <>
    <div className="product-url-intro">
      <span className="product-url-icon"><LinkIcon size={24}/></span>
      <div><strong>상품 링크만 붙여 넣어보세요</strong><p>상품명, 분류, 브랜드, 색상과 대표 이미지를 자동으로 찾아 입력해드려요.</p></div>
    </div>
    <form className="product-url-form" onSubmit={onSubmit}>
      <label><FieldTitle required>상품 URL</FieldTitle><div className="input-with-icon"><LinkIcon size={19}/><input name="productUrl" type="url" inputMode="url" autoFocus required disabled={loading} placeholder="https://www.musinsa.com/products/…"/></div></label>
      {loading && <div className="product-fetch-progress" role="status"><CircleNotch size={24}/><span><strong>상품 정보를 가져오고 있어요</strong><small>페이지의 상품 정보와 대표 이미지를 확인하는 중입니다.</small></span></div>}
      {error && <div className="product-fetch-error" role="alert"><strong>자동 입력을 완료하지 못했어요.</strong><span>{error}</span></div>}
      <p className="product-url-note">자동으로 채운 내용은 다음 화면에서 자유롭게 수정할 수 있어요. 지원되지 않는 쇼핑몰은 {subject} 정보를 직접 입력해주세요.</p>
      <footer className="modal-actions product-url-actions"><button type="button" className="secondary-button" onClick={onManual} disabled={loading}>직접 입력</button><button className="primary-button compact" disabled={loading}>{loading ? <><CircleNotch className="spin" size={18}/>처리 중</> : <><LinkIcon size={18}/>정보 가져오기</>}</button></footer>
    </form>
  </>;
}

function AutofillSummary({ product }) {
  if (!product?.autoFilled) return null;
  return <div className="autofill-summary">
    {product.image ? <img src={product.image} alt="자동으로 가져온 상품 미리보기"/> : <span className="autofill-summary-placeholder"><ImageSquare size={24}/></span>}
    <div><span><CheckCircle size={17} weight="fill"/> {product.source || "상품 페이지"}에서 가져왔어요</span><strong>내용을 확인하고 부족한 부분을 수정해주세요.</strong></div>
  </div>;
}

export function App() {
  const [theme, setTheme] = useState(() => load("mycloset-theme", "light"));
  const [view, setView] = useState(() => {
    const savedView = load("mycloset-current-view", "calendar");
    return ["calendar", "closet", "lookbook", "wishlist", "settings"].includes(savedView) ? savedView : "calendar";
  });
  useEffect(() => {
    window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }, [view]);
  const [cursor, setCursor] = useState(() => { const today = new Date(); return new Date(today.getFullYear(), today.getMonth(), 1); });
  const [selectedDate, setSelectedDate] = useState(TODAY);
  const [items, setItems] = useState(() => load("mycloset-items", INITIAL_ITEMS).map((item) => {
    const seeded = INITIAL_ITEMS.find((entry) => entry.id === item.id);
    return { ...seeded, ...item, brand: item.brand ?? seeded?.brand ?? "" };
  }));
  const [outfits, setOutfits] = useState(() => load("mycloset-outfits", INITIAL_OUTFITS));
  const [outfitNotes, setOutfitNotes] = useState(() => load("mycloset-outfit-notes", {}));
  const [wishlist, setWishlist] = useState(() => load("mycloset-wishlist", INITIAL_WISHLIST));
  const [lookbooks, setLookbooks] = useState(() => load("mycloset-lookbooks", INITIAL_LOOKBOOKS));
  const [inspirations, setInspirations] = useState(() => load("mycloset-inspirations", INITIAL_INSPIRATIONS));
  const [categories, setCategories] = useState(() => load("mycloset-categories", DEFAULT_CATEGORIES));
  const [category, setCategory] = useState("전체");
  const [wardrobeView, setWardrobeView] = useState(() => load("mycloset-wardrobe-view", "grid") === "table" ? "table" : "grid");
  const [outfitSort, setOutfitSort] = useState(() => load("mycloset-outfit-sort", "recent"));
  const [wardrobeSort, setWardrobeSort] = useState(() => load("mycloset-wardrobe-sort", "recent"));
  const [lookbookSort, setLookbookSort] = useState(() => load("mycloset-lookbook-sort", "recent"));
  const [wishlistSort, setWishlistSort] = useState(() => load("mycloset-wishlist-sort", "recent"));
  const [outfitSearch, setOutfitSearch] = useState("");
  const [wardrobeSearch, setWardrobeSearch] = useState("");
  const [lookbookSearch, setLookbookSearch] = useState("");
  const [wishlistSearch, setWishlistSearch] = useState("");
  const [outfitCategory, setOutfitCategory] = useState("전체");
  const [settingsTab, setSettingsTab] = useState(() => {
    const savedTab = load("mycloset-settings-tab", "general");
    return ["general", "closet", "backup", "reset"].includes(savedTab) ? savedTab : "general";
  });
  const [modal, setModal] = useState(null);
  const [draftOutfit, setDraftOutfit] = useState([]);
  const [draftLookbookIds, setDraftLookbookIds] = useState([]);
  const [lookbookName, setLookbookName] = useState("");
  const [lookbookMemo, setLookbookMemo] = useState("");
  const [lookbookCategory, setLookbookCategory] = useState("전체");
  const [activeLookbookId, setActiveLookbookId] = useState(null);
  const [activeInspirationId, setActiveInspirationId] = useState(null);
  const [editingLookbookId, setEditingLookbookId] = useState(null);
  const [lookbookApplyDate, setLookbookApplyDate] = useState(TODAY);
  const [lookbookCalendarCursor, setLookbookCalendarCursor] = useState(() => { const today = new Date(); return new Date(today.getFullYear(), today.getMonth(), 1); });
  const [lookbookSourceDate, setLookbookSourceDate] = useState("");
  const [editingItemId, setEditingItemId] = useState(null);
  const [itemFileName, setItemFileName] = useState("");
  const [wishFileName, setWishFileName] = useState("");
  const [productDraft, setProductDraft] = useState(null);
  const [productLookupLoading, setProductLookupLoading] = useState(false);
  const [productLookupError, setProductLookupError] = useState("");
  const [excelFileName, setExcelFileName] = useState("");
  const [backupFileName, setBackupFileName] = useState("");
  const [pendingBackup, setPendingBackup] = useState(null);
  const [pendingDeletion, setPendingDeletion] = useState(null);
  const [pinCaption, setPinCaption] = useState("");
  const [pinSourceUrl, setPinSourceUrl] = useState("");
  const [pinImage, setPinImage] = useState("");
  const [pinFileName, setPinFileName] = useState("");
  const [settingsTarget, setSettingsTarget] = useState(null);
  const [notice, setNotice] = useState("");
  const [noteEditing, setNoteEditing] = useState(false);
  const [noteDraft, setNoteDraft] = useState("");
  const [detailWidth, setDetailWidth] = useState(() => Math.max(360, Number(load("mycloset-detail-width", 440)) || 440));
  const [sidebarCollapsed, setSidebarCollapsed] = useState(() => load("mycloset-sidebar-collapsed", false));
  const [calendarOpen, setCalendarOpen] = useState(false);
  const [draggedItemId, setDraggedItemId] = useState(null);
  const [dragPreview, setDragPreview] = useState(null);
  const [reorderingItemId, setReorderingItemId] = useState(null);
  const [draggedCategoryName, setDraggedCategoryName] = useState(null);
  const [categoryDropTarget, setCategoryDropTarget] = useState(null);
  const [dropActive, setDropActive] = useState(false);
  const pointerDragRef = useRef(null);
  const suppressCardClickRef = useRef(false);
  const wardrobeDataRef = useRef(null);
  const wardrobeCollectionRef = useRef(null);

  const allCategories = useMemo(() => ["전체", ...categories], [categories]);
  useEffect(() => localStorage.setItem("mycloset-current-view", JSON.stringify(view)), [view]);
  useEffect(() => localStorage.setItem("mycloset-settings-tab", JSON.stringify(settingsTab)), [settingsTab]);
  useEffect(() => { document.documentElement.dataset.theme = theme; localStorage.setItem("mycloset-theme", JSON.stringify(theme)); }, [theme]);
  useEffect(() => localStorage.setItem("mycloset-items", JSON.stringify(items)), [items]);
  useEffect(() => localStorage.setItem("mycloset-outfits", JSON.stringify(outfits)), [outfits]);
  useEffect(() => localStorage.setItem("mycloset-outfit-notes", JSON.stringify(outfitNotes)), [outfitNotes]);
  useEffect(() => localStorage.setItem("mycloset-wishlist", JSON.stringify(wishlist)), [wishlist]);
  useEffect(() => localStorage.setItem("mycloset-lookbooks", JSON.stringify(lookbooks)), [lookbooks]);
  useEffect(() => {
    try {
      localStorage.setItem("mycloset-inspirations", JSON.stringify(inspirations));
    } catch {
      setNotice("브라우저 저장 공간이 부족해 스크랩을 저장하지 못했어요.");
    }
  }, [inspirations]);
  useEffect(() => localStorage.setItem("mycloset-categories", JSON.stringify(categories)), [categories]);
  useEffect(() => localStorage.setItem("mycloset-wardrobe-view", JSON.stringify(wardrobeView)), [wardrobeView]);
  useEffect(() => {
    wardrobeCollectionRef.current?.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }, [wardrobeView]);
  useEffect(() => localStorage.setItem("mycloset-outfit-sort", JSON.stringify(outfitSort)), [outfitSort]);
  useEffect(() => localStorage.setItem("mycloset-wardrobe-sort", JSON.stringify(wardrobeSort)), [wardrobeSort]);
  useEffect(() => localStorage.setItem("mycloset-lookbook-sort", JSON.stringify(lookbookSort)), [lookbookSort]);
  useEffect(() => localStorage.setItem("mycloset-wishlist-sort", JSON.stringify(wishlistSort)), [wishlistSort]);
  useEffect(() => localStorage.setItem("mycloset-detail-width", JSON.stringify(detailWidth)), [detailWidth]);
  useEffect(() => localStorage.setItem("mycloset-sidebar-collapsed", JSON.stringify(sidebarCollapsed)), [sidebarCollapsed]);
  useEffect(() => {
    const clampDetailWidth = () => {
      const { minWidth, maxWidth } = detailWidthLimits(sidebarCollapsed);
      setDetailWidth((current) => Math.min(maxWidth, Math.max(minWidth, current)));
    };
    clampDetailWidth();
    window.addEventListener("resize", clampDetailWidth);
    return () => window.removeEventListener("resize", clampDetailWidth);
  }, [sidebarCollapsed]);
  useEffect(() => { if (!notice) return undefined; const timer = setTimeout(() => setNotice(""), 3200); return () => clearTimeout(timer); }, [notice]);
  useEffect(() => {
    if (modal !== "inspirationCreate") return undefined;
    const pasteImage = async (event) => {
      const imageFile = [...(event.clipboardData?.files ?? [])].find((file) => file.type.startsWith("image/"));
      if (!imageFile) return;
      event.preventDefault();
      try {
        setPinImage(await optimizeScrapImage(imageFile));
        setPinFileName(imageFile.name || "붙여넣은 이미지");
      } catch (error) { setNotice(error.message || "이미지를 붙여넣지 못했어요."); }
    };
    window.addEventListener("paste", pasteImage);
    return () => window.removeEventListener("paste", pasteImage);
  }, [modal]);
  useEffect(() => {
    if (view !== "settings" || settingsTab !== "closet" || settingsTarget !== "wardrobe-data" || !wardrobeDataRef.current) return undefined;
    const frame = requestAnimationFrame(() => {
      wardrobeDataRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      wardrobeDataRef.current?.focus({ preventScroll: true });
      setSettingsTarget(null);
    });
    return () => cancelAnimationFrame(frame);
  }, [view, settingsTab, settingsTarget]);

  const itemMap = useMemo(() => Object.fromEntries(items.map((item) => [item.id, item])), [items]);
  const selectedItems = (outfits[selectedDate] ?? []).map((id) => itemMap[id]).filter(Boolean);
  const categoryFilteredItems = useMemo(() => filterItemsByCategory(items, category), [items, category]);
  const outfitVisibleItems = useMemo(() => searchWardrobeItems(categoryFilteredItems, outfitSearch), [categoryFilteredItems, outfitSearch]);
  const sortedOutfitVisibleItems = useMemo(() => sortWardrobe(outfitVisibleItems, outfitSort, categories), [outfitVisibleItems, outfitSort, categories]);
  const wardrobeVisibleItems = useMemo(() => searchWardrobeItems(categoryFilteredItems, wardrobeSearch), [categoryFilteredItems, wardrobeSearch]);
  const sortedFilteredItems = useMemo(() => sortWardrobe(wardrobeVisibleItems, wardrobeSort, categories), [wardrobeVisibleItems, wardrobeSort, categories]);
  const searchedLookbooks = useMemo(() => searchLookbooks(lookbooks, itemMap, lookbookSearch), [lookbooks, itemMap, lookbookSearch]);
  const sortedLookbooks = useMemo(() => sortLookbookEntries(searchedLookbooks, lookbookSort), [searchedLookbooks, lookbookSort]);
  const searchedWishlist = useMemo(() => searchWishlistItems(wishlist, wishlistSearch), [wishlist, wishlistSearch]);
  const sortedWishlist = useMemo(() => sortWishlistEntries(searchedWishlist, wishlistSort), [searchedWishlist, wishlistSort]);
  const outfitPickerItems = filterItemsByCategory(items, outfitCategory);
  const lookbookPickerItems = filterItemsByCategory(items, lookbookCategory);
  const activeLookbook = lookbooks.find((lookbook) => lookbook.id === activeLookbookId);
  const activeInspiration = inspirations.find((inspiration) => inspiration.id === activeInspirationId);
  const editingItem = items.find((item) => item.id === editingItemId);
  const selectedDateLabel = selectedDate.replaceAll("-", ". ");
  const selectedNote = outfitNotes[selectedDate] ?? "";
  useEffect(() => { setNoteEditing(false); setNoteDraft(outfitNotes[selectedDate] ?? ""); }, [selectedDate]);
  const handleBrandClick = () => {
    if (window.matchMedia("(max-width: 900px)").matches) { setView("calendar"); return; }
    setSidebarCollapsed((current) => !current);
  };
  const moveSelectedDate = (offset) => {
    const nextDate = new Date(`${selectedDate}T12:00:00`);
    nextDate.setDate(nextDate.getDate() + offset);
    const nextDateKey = iso(nextDate);
    setSelectedDate(nextDateKey);
    setCursor(new Date(nextDate.getFullYear(), nextDate.getMonth(), 1));
    setCalendarOpen(false);
  };
  const openWardrobeDataSettings = () => { setSettingsTarget("wardrobe-data"); setView("settings"); setSettingsTab("closet"); };
  const closeModal = () => { setModal(null); setEditingItemId(null); setActiveLookbookId(null); setActiveInspirationId(null); setEditingLookbookId(null); setItemFileName(""); setWishFileName(""); setPinCaption(""); setPinSourceUrl(""); setPinImage(""); setPinFileName(""); setPendingBackup(null); setPendingDeletion(null); setProductDraft(null); setProductLookupLoading(false); setProductLookupError(""); };
  const openOutfit = () => { setDraftOutfit([...(outfits[selectedDate] ?? [])]); setOutfitCategory("전체"); setModal("outfit"); };
  const openNewItem = () => { setEditingItemId(null); setItemFileName(""); setProductDraft(null); setProductLookupError(""); setModal("itemUrl"); };
  const openNewWish = () => { setWishFileName(""); setProductDraft(null); setProductLookupError(""); setModal("wishUrl"); };
  const openManualRegistration = (kind) => { setProductDraft(null); setProductLookupError(""); setModal(kind === "item" ? "item" : "wish"); };
  const lookupProduct = async (event, kind) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const url = String(data.get("productUrl") || "").trim();
    setProductLookupLoading(true);
    setProductLookupError("");
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), PRODUCT_LOOKUP_TIMEOUT_MS);
    try {
      const response = await fetch("/api/product-metadata", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ url }),
        signal: controller.signal,
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.error || "상품 정보를 가져오지 못했어요.");
      setProductDraft({ ...payload.product, url: payload.product?.url || url, autoFilled: true });
      setModal(kind === "item" ? "item" : "wish");
    } catch (error) {
      setProductLookupError(error?.name === "AbortError"
        ? PRODUCT_LOOKUP_TIMEOUT_MESSAGE
        : error instanceof Error ? error.message : "상품 정보를 가져오지 못했어요.");
    } finally {
      window.clearTimeout(timeoutId);
      setProductLookupLoading(false);
    }
  };
  const openEditItem = (id) => { setEditingItemId(id); setItemFileName(""); setModal("editItem"); };
  const toggleDraft = (id) => setDraftOutfit((current) => current.includes(id) ? current.filter((itemId) => itemId !== id) : [...current, id]);
  const openLookbookCreator = ({ itemIds = [], sourceDate = "" } = {}) => {
    setEditingLookbookId(null);
    setDraftLookbookIds([...itemIds]);
    setLookbookName(sourceDate ? `${sourceDate.replaceAll("-", ". ")} 착장` : nextLookbookName(lookbooks));
    setLookbookMemo(sourceDate ? (outfitNotes[sourceDate] ?? "") : "");
    setLookbookSourceDate(sourceDate);
    setLookbookCategory("전체");
    setModal("lookbookCreate");
  };
  const openLookbookEditor = (id) => {
    const lookbook = lookbooks.find((entry) => entry.id === id);
    if (!lookbook) return;
    setEditingLookbookId(id);
    setDraftLookbookIds([...lookbook.itemIds]);
    setLookbookName(lookbook.name);
    setLookbookMemo(lookbook.memo || "");
    setLookbookSourceDate(lookbook.sourceDate || "");
    setLookbookCategory("전체");
    setModal("lookbookCreate");
  };
  const toggleLookbookDraft = (id) => setDraftLookbookIds((current) => current.includes(id) ? current.filter((itemId) => itemId !== id) : [...current, id]);
  const saveLookbook = () => {
    const name = lookbookName.trim();
    if (!name) { setNotice("룩북 이름을 입력해주세요."); return; }
    if (!draftLookbookIds.length) { setNotice("룩북에 한 가지 이상의 옷을 담아주세요."); return; }
    setLookbooks((current) => editingLookbookId
      ? current.map((lookbook) => lookbook.id === editingLookbookId
        ? { ...lookbook, name, memo: lookbookMemo.trim(), itemIds: [...draftLookbookIds] }
        : lookbook)
      : [...current, { id: `lookbook-${Date.now()}`, name, memo: lookbookMemo.trim(), itemIds: [...draftLookbookIds], sourceDate: lookbookSourceDate }]);
    closeModal();
    setNotice(editingLookbookId ? "룩북을 수정했어요." : "룩북에 새 착장을 저장했어요.");
  };
  const openLookbookApply = (id) => {
    const initialDate = selectedDate || TODAY;
    const initialCursor = new Date(`${initialDate}T00:00:00`);
    setActiveLookbookId(id);
    setLookbookApplyDate(initialDate);
    setLookbookCalendarCursor(new Date(initialCursor.getFullYear(), initialCursor.getMonth(), 1));
    setModal("lookbookApply");
  };
  const applyLookbook = () => {
    const targetDate = lookbookApplyDate;
    if (!activeLookbook || !targetDate) return;
    const validIds = validItemIds(activeLookbook.itemIds, itemMap);
    setOutfits((current) => ({ ...current, [targetDate]: validIds }));
    const date = new Date(`${targetDate}T00:00:00`);
    setSelectedDate(targetDate);
    setCursor(new Date(date.getFullYear(), date.getMonth(), 1));
    setView("calendar");
    closeModal();
    setNotice(`${activeLookbook.name}을(를) ${targetDate.replaceAll("-", ". ")}에 적용했어요.`);
  };
  const openLookbookDelete = (lookbook) => {
    setPendingDeletion({ type: "lookbook", id: lookbook.id, name: lookbook.name });
    setModal("confirmLookbook");
  };
  const deleteLookbook = () => {
    if (pendingDeletion?.type !== "lookbook") return;
    setLookbooks((current) => current.filter((lookbook) => lookbook.id !== pendingDeletion.id));
    closeModal();
    setNotice("룩북을 삭제했어요.");
  };
  const openInspirationCreator = () => {
    setPinCaption("");
    setPinSourceUrl("");
    setPinImage("");
    setPinFileName("");
    setModal("inspirationCreate");
  };
  const openInspirationViewer = (pin) => {
    setActiveInspirationId(pin.id);
    setModal("inspirationViewer");
  };
  const handlePinFile = async (file) => {
    if (!file) return;
    try {
      setPinImage(await optimizeScrapImage(file));
      setPinFileName(file.name);
    } catch (error) { setNotice(error.message || "이미지를 불러오지 못했어요."); }
  };
  const saveInspiration = (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const imageUrl = String(data.get("imageUrl") || "").trim();
    const image = pinImage || imageUrl;
    if (!image) { setNotice("코디 이미지 URL을 입력하거나 이미지를 붙여넣어 주세요."); return; }
    const nextInspiration = {
      id: `inspiration-${Date.now()}`,
      image,
      caption: pinCaption.trim(),
      sourceUrl: pinSourceUrl.trim(),
      createdAt: new Date().toISOString(),
    };
    const nextInspirations = [nextInspiration, ...inspirations];
    try {
      localStorage.setItem("mycloset-inspirations", JSON.stringify(nextInspirations));
    } catch {
      setNotice("브라우저 저장 공간이 부족해 스크랩을 저장하지 못했어요. 큰 스크랩을 삭제한 뒤 다시 시도해주세요.");
      return;
    }
    setInspirations(nextInspirations);
    closeModal();
    setNotice("마음에 든 코디를 스크랩했어요.");
  };
  const openInspirationDelete = (pin) => {
    setPendingDeletion({ type: "inspiration", id: pin.id, name: pin.caption || "저장한 코디" });
    setModal("confirmInspiration");
  };
  const deleteInspiration = () => {
    if (pendingDeletion?.type !== "inspiration") return;
    setInspirations((current) => current.filter((pin) => pin.id !== pendingDeletion.id));
    closeModal();
    setNotice("스크랩을 삭제했어요.");
  };
  const openWishlistDelete = (item) => {
    setPendingDeletion({ type: "wishlist", id: item.id, name: item.name });
    setModal("confirmWishlist");
  };
  const deleteWishlistItem = () => {
    if (pendingDeletion?.type !== "wishlist") return;
    setWishlist((current) => current.filter((item) => item.id !== pendingDeletion.id));
    closeModal();
    setNotice("위시리스트에서 삭제했어요.");
  };
  const addItemToSelectedOutfit = (id) => {
    if (!itemMap[id]) return;
    const current = outfits[selectedDate] ?? [];
    if (current.includes(id)) { setNotice("이미 오늘의 착장에 있는 옷이에요."); return; }
    setOutfits((existing) => ({ ...existing, [selectedDate]: [...(existing[selectedDate] ?? []), id] }));
    setNotice(`${itemMap[id].name}을(를) ${selectedDateLabel} 착장에 추가했어요.`);
  };
  const removeItemFromSelectedOutfit = (id) => {
    setOutfits((current) => ({ ...current, [selectedDate]: (current[selectedDate] ?? []).filter((itemId) => itemId !== id) }));
    setNotice(`${itemMap[id]?.name ?? "옷"}을(를) 착장에서 뺐어요.`);
  };
  const reorderSelectedOutfit = (sourceId, targetId) => {
    if (!sourceId || sourceId === targetId) return;
    setOutfits((current) => {
      const list = current[selectedDate] ?? [];
      const next = reorderItemIds(list, sourceId, targetId);
      return next === list ? current : { ...current, [selectedDate]: next };
    });
  };
  const moveSelectedItem = (id, direction) => {
    setOutfits((current) => {
      const list = current[selectedDate] ?? [];
      const next = moveItemByOffset(list, id, direction);
      return next === list ? current : { ...current, [selectedDate]: next };
    });
  };
  const startOutfitReorder = (event, id) => {
    if (event.pointerType !== "touch" && event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const drag = { id, startX: event.clientX, startY: event.clientY, active: false, lastTarget: id };
    const onMove = (moveEvent) => {
      if (!drag.active && Math.hypot(moveEvent.clientX - drag.startX, moveEvent.clientY - drag.startY) < 6) return;
      if (!drag.active) { drag.active = true; setReorderingItemId(id); }
      const targetId = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY)?.closest(".outfit-showcase-card")?.dataset.itemId;
      if (targetId && targetId !== drag.lastTarget) {
        reorderSelectedOutfit(id, targetId);
        drag.lastTarget = targetId;
      }
    };
    const onUp = () => {
      if (drag.active) setNotice("착장 순서를 변경했어요.");
      setReorderingItemId(null);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
  };
  const startItemPointerDrag = (event, id) => {
    if (event.pointerType === "mouse" || (event.pointerType === "pen" && event.button !== 0)) return;
    pointerDragRef.current = { id, pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, active: false };
    event.currentTarget.setPointerCapture?.(event.pointerId);
  };
  const moveItemPointerDrag = (event) => {
    const drag = pointerDragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (!drag.active && Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 8) return;
    if (!drag.active) { drag.active = true; setDraggedItemId(drag.id); }
    setDragPreview({ id: drag.id, x: event.clientX, y: event.clientY });
    const overOutfit = Boolean(document.elementFromPoint(event.clientX, event.clientY)?.closest(".today-outfit"));
    setDropActive(overOutfit);
  };
  const finishItemPointerDrag = (event) => {
    const drag = pointerDragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (drag.active) {
      suppressCardClickRef.current = true;
      const overOutfit = Boolean(document.elementFromPoint(event.clientX, event.clientY)?.closest(".today-outfit"));
      if (overOutfit) addItemToSelectedOutfit(drag.id);
    }
    pointerDragRef.current = null;
    setDraggedItemId(null);
    setDragPreview(null);
    setDropActive(false);
  };
  const startItemMouseDrag = (event, id) => {
    if (event.button !== 0) return;
    const drag = { id, startX: event.clientX, startY: event.clientY, active: false };
    const onMove = (moveEvent) => {
      if (!drag.active && Math.hypot(moveEvent.clientX - drag.startX, moveEvent.clientY - drag.startY) < 8) return;
      if (!drag.active) { drag.active = true; setDraggedItemId(id); }
      setDragPreview({ id, x: moveEvent.clientX, y: moveEvent.clientY });
      setDropActive(Boolean(document.elementFromPoint(moveEvent.clientX, moveEvent.clientY)?.closest(".today-outfit")));
    };
    const onUp = (upEvent) => {
      if (drag.active) {
        suppressCardClickRef.current = true;
        if (document.elementFromPoint(upEvent.clientX, upEvent.clientY)?.closest(".today-outfit")) addItemToSelectedOutfit(id);
      }
      setDraggedItemId(null);
      setDragPreview(null);
      setDropActive(false);
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };
  const saveOutfit = () => { setOutfits((current) => ({ ...current, [selectedDate]: draftOutfit })); closeModal(); setNotice("착장을 저장했어요."); };
  const beginNoteEditing = () => { setNoteDraft(selectedNote); setNoteEditing(true); };
  const cancelNoteEditing = () => { setNoteDraft(selectedNote); setNoteEditing(false); };
  const confirmNoteSave = () => {
    const nextNote = noteDraft.trim();
    setOutfitNotes((current) => ({ ...current, [selectedDate]: nextNote }));
    setNoteDraft(nextNote);
    setNoteEditing(false);
    setNotice("메모가 저장되었습니다.");
  };
  const deleteOutfit = () => {
    setOutfits((current) => { const next = { ...current }; delete next[selectedDate]; return next; });
    setOutfitNotes((current) => { const next = { ...current }; delete next[selectedDate]; return next; });
    closeModal(); setNotice("이 날짜의 착장을 삭제했어요.");
  };
  const resetWardrobeData = () => {
    const reset = createWardrobeResetState(DEFAULT_CATEGORIES);
    setItems(reset.items);
    setOutfits(reset.outfits);
    setOutfitNotes(reset.outfitNotes);
    setLookbooks(reset.lookbooks);
    setCategories(reset.categories);
    setCategory("전체");
    setOutfitCategory("전체");
    setLookbookCategory("전체");
    setNoteDraft("");
    setNoteEditing(false);
    closeModal();
    setNotice("옷장과 연결된 데이터를 초기화했어요.");
  };
  const startRailResize = (event) => {
    if (window.innerWidth <= 900) return;
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = detailWidth;
    const onMove = (moveEvent) => {
      const { minWidth, maxWidth } = detailWidthLimits(sidebarCollapsed);
      setDetailWidth(Math.min(maxWidth, Math.max(minWidth, startWidth - (moveEvent.clientX - startX) / DESKTOP_UI_DENSITY)));
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      document.body.classList.remove("is-resizing");
    };
    document.body.classList.add("is-resizing");
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };
  const resizeRailWithKeyboard = (event) => {
    if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
    event.preventDefault();
    const { minWidth, maxWidth } = detailWidthLimits(sidebarCollapsed);
    setDetailWidth((current) => Math.min(maxWidth, Math.max(minWidth, current + (event.key === "ArrowLeft" ? 24 : -24) / DESKTOP_UI_DENSITY)));
  };

  const itemFromForm = async (event, previous = {}) => {
    const data = new FormData(event.currentTarget);
    const file = data.get("imageFile");
    let image = String(data.get("image") || "").trim();
    if (file instanceof File && file.size) image = await fileToDataUrl(file);
    return {
      ...previous,
      name: String(data.get("name") || "").trim(),
      brand: String(data.get("brand") || "").trim(),
      category: String(data.get("category") || "").trim(),
      color: String(data.get("color") || "").trim(),
      url: String(data.get("url") || "").trim(),
      image,
    };
  };
  const addItem = async (event) => {
    event.preventDefault();
    try {
      const nextItem = await itemFromForm(event, { id: `item-${Date.now()}` });
      setItems((current) => [...current, nextItem]); closeModal(); setNotice("새 옷을 옷장에 추가했어요.");
    } catch (error) { setNotice(error.message); }
  };
  const updateItem = async (event) => {
    event.preventDefault();
    try {
      const nextItem = await itemFromForm(event, editingItem);
      setItems((current) => current.map((item) => item.id === editingItemId ? nextItem : item)); closeModal(); setNotice("옷 정보를 수정했어요.");
    } catch (error) { setNotice(error.message); }
  };
  const deleteItem = () => {
    setItems((current) => current.filter((item) => item.id !== editingItemId));
    setOutfits((current) => removeItemReferences(current, [], editingItemId).outfits);
    setLookbooks((current) => removeItemReferences({}, current, editingItemId).lookbooks);
    closeModal(); setNotice("옷장에서 삭제했어요.");
  };
  const addWish = async (event) => {
    event.preventDefault();
    try {
      const data = new FormData(event.currentTarget);
      const file = data.get("imageFile");
      let image = String(data.get("image") || "").trim();
      if (file instanceof File && file.size) image = await fileToDataUrl(file);
      setWishlist((current) => [...current, {
        id: `wish-${Date.now()}`, name: String(data.get("name") || "").trim(), brand: String(data.get("brand") || "").trim(),
        category: String(data.get("category") || "").trim(), color: String(data.get("color") || "").trim(), status: "관심", url: String(data.get("url") || "").trim(), image,
      }]);
      closeModal(); setNotice("위시리스트에 스크랩했어요.");
    } catch (error) { setNotice(error.message); }
  };

  const addCategory = (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const name = String(data.get("categoryName") || "").trim();
    if (!name || name === "전체") return;
    if (categories.includes(name)) { setNotice("이미 있는 카테고리예요."); return; }
    setCategories((current) => [...current, name]); event.currentTarget.reset(); setNotice(`${name} 카테고리를 추가했어요.`);
  };
  const deleteCategory = (name) => {
    if (items.some((item) => item.category === name)) { setNotice("사용 중인 카테고리는 옷을 다른 분류로 옮긴 뒤 삭제할 수 있어요."); return; }
    setCategories((current) => current.filter((categoryName) => categoryName !== name));
    if (category === name) setCategory("전체");
    if (outfitCategory === name) setOutfitCategory("전체");
    setNotice(`${name} 카테고리를 삭제했어요.`);
  };
  const moveCategoryToIndex = (name, nextIndex) => {
    setCategories((current) => {
      const currentIndex = current.indexOf(name);
      if (currentIndex < 0) return current;
      const boundedIndex = Math.max(0, Math.min(current.length - 1, nextIndex));
      if (currentIndex === boundedIndex) return current;
      const next = [...current];
      next.splice(currentIndex, 1);
      next.splice(boundedIndex, 0, name);
      return next;
    });
  };
  const moveCategory = (name, direction) => {
    const currentIndex = categories.indexOf(name);
    const nextIndex = currentIndex + direction;
    if (currentIndex < 0 || nextIndex < 0 || nextIndex >= categories.length) return;
    moveCategoryToIndex(name, nextIndex);
    setNotice(`${name} 카테고리 순서를 변경했어요.`);
  };
  const handleCategoryKeyDown = (event, name) => {
    const currentIndex = categories.indexOf(name);
    if (event.key === "ArrowUp") { event.preventDefault(); moveCategory(name, -1); }
    if (event.key === "ArrowDown") { event.preventDefault(); moveCategory(name, 1); }
    if (event.key === "Home" && currentIndex > 0) { event.preventDefault(); moveCategoryToIndex(name, 0); setNotice(`${name} 카테고리를 맨 앞으로 옮겼어요.`); }
    if (event.key === "End" && currentIndex < categories.length - 1) { event.preventDefault(); moveCategoryToIndex(name, categories.length - 1); setNotice(`${name} 카테고리를 맨 뒤로 옮겼어요.`); }
  };
  const startCategoryDrag = (event, name) => {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", name);
    setDraggedCategoryName(name);
    setCategoryDropTarget(name);
  };
  const finishCategoryDrop = (event, targetName) => {
    event.preventDefault();
    const sourceName = draggedCategoryName || event.dataTransfer.getData("text/plain");
    if (sourceName && sourceName !== targetName) {
      moveCategoryToIndex(sourceName, categories.indexOf(targetName));
      setNotice(`${sourceName} 카테고리 순서를 변경했어요.`);
    }
    setDraggedCategoryName(null);
    setCategoryDropTarget(null);
  };
  const importExcel = async (file) => {
    if (!file) return;
    setExcelFileName(file.name);
    try {
      const XLSX = await import("xlsx");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames.includes("옷장 업로드") ? "옷장 업로드" : workbook.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      const imported = parseWardrobeExcelRows(rows);
      if (!imported.length) throw new Error("등록할 옷을 찾지 못했어요. 이름과 분류 열을 확인해 주세요.");
      const importedCategories = [...new Set(imported.map((item) => item.category).filter(Boolean))];
      setCategories((current) => [...new Set([...current, ...importedCategories])]);
      setItems((current) => [...current, ...imported]);
      setNotice(`${imported.length}개의 옷을 한 번에 추가했어요.`);
    } catch (error) { setNotice(error.message || "Excel 파일을 읽지 못했어요."); }
  };

  const exportExcel = async () => {
    if (!items.length) { setNotice("등록된 옷이 없어 내보낼 수 없어요. 먼저 옷을 등록하거나 Excel로 가져와 주세요."); return; }
    try {
      const XLSX = await import("xlsx");
      const rows = [WARDROBE_EXCEL_HEADERS, ...wardrobeItemsToExcelRows(items)];
      const sheet = XLSX.utils.aoa_to_sheet(rows);
      sheet["!cols"] = [{ wch: 28 }, { wch: 15 }, { wch: 20 }, { wch: 18 }, { wch: 46 }, { wch: 46 }];
      sheet["!autofilter"] = { ref: `A1:F${rows.length}` };
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, "옷장 업로드");
      XLSX.writeFile(workbook, `MyCloset_옷장_${TODAY}.xlsx`);
      setNotice(`${items.length}개의 옷장 데이터를 Excel로 내보냈어요.`);
    } catch (error) { setNotice(error.message || "Excel 파일을 만들지 못했어요."); }
  };

  const exportBackup = () => {
    const payload = createBackupPayload(
      { items, categories, outfits, outfitNotes, lookbooks, wishlist, inspirations },
      {
        theme, sidebarCollapsed, detailWidth, currentView: view, settingsTab,
        wardrobeView, outfitSort, wardrobeSort, lookbookSort, wishlistSort,
      },
    );
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `mycloset-backup-${formatFileTimestamp()}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setNotice("전체 데이터 백업 파일을 저장했어요.");
  };
  const prepareBackupImport = async (file) => {
    if (!file) return;
    setBackupFileName(file.name);
    try {
      if (file.size > 25 * 1024 * 1024) throw new Error("백업 파일은 25MB 이하만 가져올 수 있어요.");
      const validated = validateBackupPayload(JSON.parse(await file.text()));
      setPendingBackup(validated);
      setModal("confirmBackupImport");
    } catch (error) {
      setPendingBackup(null);
      setNotice(error instanceof SyntaxError ? "JSON 백업 파일을 읽지 못했어요." : error.message);
    }
  };
  const importBackup = () => {
    if (!pendingBackup) return;
    const { data, preferences } = pendingBackup;
    try {
      writeLocalStorageSnapshot([
        ["mycloset-items", data.items],
        ["mycloset-categories", data.categories],
        ["mycloset-outfits", data.outfits],
        ["mycloset-outfit-notes", data.outfitNotes],
        ["mycloset-lookbooks", data.lookbooks],
        ["mycloset-wishlist", data.wishlist],
        ["mycloset-inspirations", data.inspirations ?? []],
      ]);
    } catch {
      setNotice("브라우저 저장 공간이 부족해 백업을 가져오지 못했어요. 현재 데이터는 유지됐어요.");
      return;
    }
    setItems(data.items);
    setCategories(data.categories);
    setOutfits(data.outfits);
    setOutfitNotes(data.outfitNotes);
    setLookbooks(data.lookbooks);
    setWishlist(data.wishlist);
    setInspirations(data.inspirations ?? []);
    if (preferences.theme === "light" || preferences.theme === "dark") setTheme(preferences.theme);
    if (typeof preferences.sidebarCollapsed === "boolean") setSidebarCollapsed(preferences.sidebarCollapsed);
    if (["calendar", "closet", "lookbook", "wishlist", "settings"].includes(preferences.currentView)) setView(preferences.currentView);
    if (["general", "closet", "backup", "reset"].includes(preferences.settingsTab)) setSettingsTab(preferences.settingsTab);
    if (preferences.wardrobeView === "grid" || preferences.wardrobeView === "table") setWardrobeView(preferences.wardrobeView);
    if (["recent", "oldest", "name", "category", "brand"].includes(preferences.outfitSort)) setOutfitSort(preferences.outfitSort);
    if (["recent", "oldest", "name", "category", "brand"].includes(preferences.wardrobeSort)) setWardrobeSort(preferences.wardrobeSort);
    if (["recent", "oldest", "name", "items"].includes(preferences.lookbookSort)) setLookbookSort(preferences.lookbookSort);
    if (["recent", "oldest", "name", "brand"].includes(preferences.wishlistSort)) setWishlistSort(preferences.wishlistSort);
    if (Number.isFinite(Number(preferences.detailWidth))) {
      const collapsed = typeof preferences.sidebarCollapsed === "boolean" ? preferences.sidebarCollapsed : sidebarCollapsed;
      const { minWidth, maxWidth } = detailWidthLimits(collapsed);
      setDetailWidth(Math.min(maxWidth, Math.max(minWidth, Number(preferences.detailWidth))));
    }
    setCategory("전체");
    setOutfitCategory("전체");
    setLookbookCategory("전체");
    setNoteEditing(false);
    setNoteDraft("");
    setBackupFileName("");
    closeModal();
    setNotice("전체 데이터를 가져왔어요.");
  };

  const renderItemForm = ({ mode }) => {
    const isEdit = mode === "edit";
    const item = isEdit ? editingItem : (productDraft || {});
    return <><AutofillSummary product={isEdit ? null : productDraft}/><form className="entry-form" onSubmit={isEdit ? updateItem : addItem}>
      <label><FieldTitle required>이름</FieldTitle><input name="name" required defaultValue={item?.name || ""} placeholder="예: 네이비 가디건"/></label>
      <div className="form-row"><label><FieldTitle required>분류</FieldTitle><SelectField name="category" required defaultValue={item?.category || categories[0]}>{categories.map((name) => <option key={name}>{name}</option>)}</SelectField></label><label><FieldTitle>브랜드</FieldTitle><input name="brand" defaultValue={item?.brand || ""} placeholder="예: COS"/></label></div>
      <label><FieldTitle>색상</FieldTitle><input name="color" defaultValue={item?.color || ""} placeholder="예: 네이비"/></label>
      <label><FieldTitle>상품 URL</FieldTitle><div className="input-with-icon"><LinkIcon size={19}/><input name="url" type="url" defaultValue={item?.url || ""} placeholder="https://"/></div></label>
      <div className="field-group image-source-field"><FieldTitle>이미지 URL 또는 파일</FieldTitle><div className="image-source-row"><div className="input-with-icon"><ImageSquare size={19}/><input name="image" type="text" inputMode="url" defaultValue={item?.image || ""} placeholder="https://image.example.com/item.jpg"/></div><input id={`${mode}-item-image`} className="visually-hidden-input" name="imageFile" type="file" accept="image/*" onChange={(event) => setItemFileName(event.target.files?.[0]?.name || "")}/><label className="image-upload-button" htmlFor={`${mode}-item-image`} title="이미지 파일 선택"><UploadSimple size={18}/><span>파일</span></label></div><small className={`field-help image-source-help${itemFileName ? " selected" : ""}`} aria-live="polite">{itemFileName ? `${itemFileName} 선택됨 · URL 이미지 대신 이 파일을 사용해요.` : "이미지 URL을 입력하거나 JPG, PNG, WEBP 파일을 선택하세요. 최대 2.5MB"}</small></div>
      <footer className="modal-actions">{isEdit && <button type="button" className="danger-button" onClick={() => setModal("confirmItem")}><Trash size={18}/>삭제</button>}<span className="action-spacer"/><button type="button" className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact">{isEdit ? "수정 저장" : "옷장에 추가"}</button></footer>
    </form></>;
  };

  return <div className={`app-shell${sidebarCollapsed ? " sidebar-collapsed" : ""}`}>
    <aside className="sidebar">
      <div className="sidebar-top"><button className="brand" onClick={handleBrandClick} aria-label={sidebarCollapsed ? "MyCloset 사이드바 펼치기" : "MyCloset 사이드바 접기"} aria-pressed={sidebarCollapsed} title={sidebarCollapsed ? "사이드바 펼치기" : "사이드바 접기"}><StackSimple className="brand-mark" size={25}/><span className="brand-word">MyCloset</span></button></div>
      <nav className="main-nav">
        <button className={view === "calendar" ? "active" : ""} onClick={() => setView("calendar")} aria-label="착장" title="착장"><CalendarDots size={22}/><span>착장</span></button>
        <button className={view === "closet" ? "active" : ""} onClick={() => setView("closet")} aria-label="옷장" title="옷장"><Hanger size={22}/><span>옷장</span></button>
        <button className={view === "lookbook" ? "active" : ""} onClick={() => setView("lookbook")} aria-label="룩북" title="룩북"><BookOpen size={22}/><span>룩북</span></button>
        <button className={view === "wishlist" ? "active" : ""} onClick={() => setView("wishlist")} aria-label="위시리스트" title="위시리스트"><Heart size={22}/><span>위시리스트</span></button>
      </nav>
      <div className="sidebar-footer"><button className={view === "settings" ? "active" : ""} onClick={() => setView("settings")} aria-label="설정" title="설정"><GearSix size={21}/><span>설정</span></button><button onClick={() => setTheme(theme === "light" ? "dark" : "light")} aria-label={theme === "light" ? "다크 모드" : "라이트 모드"} title={theme === "light" ? "다크 모드" : "라이트 모드"}>{theme === "light" ? <Moon size={21}/> : <Sun size={21}/>}<span>{theme === "light" ? "다크 모드" : "라이트 모드"}</span></button></div>
    </aside>

    {view === "calendar" && <main className="calendar-page closet-first-page" style={{ "--detail-width": `${detailWidth}px`, "--dense-detail-width": `${Math.round(detailWidth * DESKTOP_UI_DENSITY)}px` }}>
      <section className="closet-rail closet-main-panel">
        <div className="closet-main-controls">
          <div className="rail-title-row closet-main-heading"><div><span className="mobile-section-label">STEP 2</span><h1>내 옷장</h1><p>{items.length ? "카드를 눌러 오늘의 착장에 추가해보세요." : "첫 옷을 등록하고 오늘의 착장을 만들어보세요."}</p></div>{items.length > 0 && <div className="closet-main-actions"><button className="secondary-button compact" onClick={() => setView("closet")}><PencilSimple size={18}/>옷장 관리</button><button className="primary-button compact" onClick={openNewItem}><Plus size={18}/>새 옷 등록</button></div>}</div>
          {items.length > 0 && <div className="closet-main-filter-row"><div className="category-tabs rail-category-tabs closet-main-tabs" aria-label="옷장 카테고리">{allCategories.map((name) => <button className={category === name ? "active" : ""} onClick={() => setCategory(name)} key={name}>{name}</button>)}</div><div className="closet-main-filter-actions"><SearchField value={outfitSearch} onChange={setOutfitSearch} label="착장에 추가할 옷 검색" placeholder="이름, 브랜드, 색상 검색"/><SortControl label="착장용 옷 정렬" value={outfitSort} onChange={setOutfitSort}><option value="recent">최근 등록순</option><option value="oldest">오래된 등록순</option><option value="name">이름순</option><option value="category">카테고리순</option><option value="brand">브랜드순</option></SortControl></div></div>}
        </div>
        {items.length ? (sortedOutfitVisibleItems.length ? <div className="closet-mini-grid">{sortedOutfitVisibleItems.map((item) => { const isInOutfit = (outfits[selectedDate] ?? []).includes(item.id); return <button className={`${draggedItemId === item.id ? "is-dragging" : ""}${isInOutfit ? " is-in-outfit" : ""}`} key={item.id} onMouseDown={(event) => startItemMouseDrag(event, item.id)} onPointerDown={(event) => startItemPointerDrag(event, item.id)} onPointerMove={moveItemPointerDrag} onPointerUp={finishItemPointerDrag} onPointerCancel={finishItemPointerDrag} onClick={() => { if (suppressCardClickRef.current) { suppressCardClickRef.current = false; return; } addItemToSelectedOutfit(item.id); }} aria-label={`${item.name}, ${isInOutfit ? "선택한 날짜 착장에 추가됨" : `${selectedDateLabel} 착장에 바로 추가`}`} aria-pressed={isInOutfit} title={isInOutfit ? "이미 선택한 날짜의 착장에 있어요" : `${selectedDateLabel} 착장에 추가`}><ItemVisual item={item}/><span className="item-meta"><strong>{item.name}</strong><small>{item.brand || "브랜드 미지정"}</small><small>{item.category} · {item.color || "색상 미지정"}</small></span>{isInOutfit && <span className="outfit-added-badge" aria-hidden="true"><Check size={13} weight="bold"/></span>}</button>; })}</div> : <SearchEmptyState subject="옷" query={outfitSearch} category={category} onClear={() => setOutfitSearch("")} onShowAll={() => setCategory("전체")}/>) : <section className="closet-empty-state"><span className="empty-state-icon"><Hanger size={28} weight="light"/></span><div className="empty-state-copy"><h2>내 옷장을 시작해볼까요?</h2><p>상품 URL을 붙여 넣으면 이름, 브랜드, 색상과 이미지를 자동으로 채워드려요.</p></div><div className="empty-state-actions"><button className="primary-button compact" onClick={openNewItem}><Plus size={18}/>첫 옷 등록</button><button className="secondary-button compact" onClick={openWardrobeDataSettings}><FileXls size={18}/>Excel로 한 번에 등록</button></div></section>}
        {items.length > 0 && <button className="mobile-closet-more" type="button" onClick={() => setView("closet")}>옷장 전체 보기 <CaretRight size={16} weight="bold"/></button>}
      </section>
      <section className="calendar-main outfit-main">
        <header className="outfit-main-header"><div className="outfit-heading-copy"><span className="mobile-section-label">STEP 1</span><h1>오늘의 착장</h1><p>{selectedDateLabel}</p><small className="mobile-outfit-help">{items.length ? "옷장에서 아이템을 골라 오늘 입을 조합을 만들어보세요." : "먼저 옷 한 벌을 등록하면 착장을 시작할 수 있어요."}</small></div><div className="outfit-date-quick-nav" role="group" aria-label="착장 날짜 빠른 이동"><button type="button" onClick={() => moveSelectedDate(-1)} aria-label="이전 날짜로 이동" title="이전 날짜"><CaretLeft size={18}/></button><button type="button" onClick={() => moveSelectedDate(1)} aria-label="다음 날짜로 이동" title="다음 날짜"><CaretRight size={18}/></button></div></header>
        <section className={`today-outfit outfit-workspace${dropActive ? " drop-active" : ""}${items.length ? "" : " is-empty-wardrobe"}`} aria-label="오늘의 착장 드롭 영역">
          {selectedItems.length ? <div className="outfit-showcase">{selectedItems.map((item) => <article className={`outfit-showcase-card${reorderingItemId === item.id ? " is-reordering" : ""}`} data-item-id={item.id} key={item.id} tabIndex="0" onPointerDown={(event) => { if (event.target.closest("button,a")) return; startOutfitReorder(event, item.id); }} onKeyDown={(event) => { if (event.key === "ArrowLeft") moveSelectedItem(item.id, -1); if (event.key === "ArrowRight") moveSelectedItem(item.id, 1); }} aria-label={`${item.name}. 카드를 드래그하거나 좌우 화살표로 순서 변경`} title="카드를 드래그해서 순서 변경">
            <div className="outfit-card-tools">{item.url && <a className="card-product-link" href={item.url} target="_blank" rel="noreferrer" aria-label={`${item.name} 상품 페이지 열기`} title="상품 페이지 열기"><LinkIcon size={17}/></a>}<button type="button" className="card-remove" onClick={() => removeItemFromSelectedOutfit(item.id)} aria-label={`${item.name} 착장에서 제거`} title="착장에서 제거"><X size={16} weight="bold"/></button></div>
            <div className="outfit-showcase-visual"><ItemVisual item={item}/></div><div className="outfit-showcase-meta"><strong>{item.name}</strong><span>{item.brand || "브랜드 미지정"}</span><small>{item.category} · {item.color || "색상 미지정"}</small></div></article>)}</div> : <div className="outfit-workspace-empty"><Hanger size={34} weight="light"/><strong>{items.length ? "오늘 입을 옷을 골라보세요" : "옷을 등록하면 착장을 만들 수 있어요"}</strong><span>{items.length ? "내 옷장에서 아이템을 선택해 조합을 시작하세요." : "첫 옷 등록은 내 옷장에서 시작할 수 있어요."}</span>{!items.length && <button className="primary-button compact mobile-empty-primary" onClick={openNewItem}><Plus size={18}/>첫 옷 등록</button>}</div>}
          {items.length > 0 && <footer className="outfit-workspace-actions"><div><strong>{selectedItems.length ? `${selectedItems.length}개의 아이템` : "착장을 시작해보세요"}</strong>{!selectedItems.length && <span>여러 아이템을 골라 나만의 조합을 만들 수 있어요.</span>}</div><div className="outfit-action-buttons">{selectedItems.length > 0 && <button className="secondary-button compact" onClick={() => openLookbookCreator({ itemIds: outfits[selectedDate] ?? [], sourceDate: selectedDate })}><BookmarkSimple size={18}/>룩북 저장</button>}<button className="primary-button compact" onClick={openOutfit}>{selectedItems.length ? <PencilSimple size={18}/> : <Plus size={18}/>} {selectedItems.length ? "착장 수정" : "착장 만들기"}</button></div></footer>}
        </section>
        <section className={`outfit-note-card${noteEditing ? " is-editing" : ""}`} aria-label={`${selectedDateLabel} 메모`}>
          <div className="outfit-note-summary"><span className="outfit-note-icon"><NotePencil size={19}/></span><div className="outfit-note-copy"><div><strong>메모</strong>{selectedNote && !noteEditing && <small><CheckCircle size={13} weight="fill"/>저장됨</small>}</div>{!noteEditing && <p className={selectedNote ? "" : "is-empty"}>{selectedNote || "이 날짜에 기억할 내용을 남겨보세요."}</p>}</div>{!noteEditing && <button type="button" className="outfit-note-edit" onClick={beginNoteEditing}><PencilSimple size={16}/>{selectedNote ? "수정" : "메모 작성"}</button>}</div>
          {noteEditing && <div className="outfit-note-editor"><textarea autoFocus maxLength="240" value={noteDraft} onChange={(event) => setNoteDraft(event.target.value)} onKeyDown={(event) => { if ((event.metaKey || event.ctrlKey) && event.key === "Enter") confirmNoteSave(); if (event.key === "Escape") cancelNoteEditing(); }} aria-label="메모 내용" placeholder="약속, 날씨, 준비물처럼 이 날짜에 기억할 내용을 적어보세요."/><footer><span>{noteDraft.length}/240 · ⌘ Enter로 저장</span><div><button type="button" className="outfit-note-cancel" onClick={cancelNoteEditing}>취소</button><button type="button" className="outfit-note-save" onClick={confirmNoteSave}><Check size={16} weight="bold"/>메모 저장</button></div></footer></div>}
        </section>
      </section>
      <button className="rail-resizer" type="button" onPointerDown={startRailResize} onKeyDown={resizeRailWithKeyboard} aria-label="내 옷장과 보조 영역 너비 조절" title="드래그해서 영역 너비 조절"/>
      <aside className="detail-rail calendar-detail-rail closet-first-rail">
        {calendarOpen && <CompactCalendar className="rail-calendar" id="outfit-calendar-popover" role="dialog" ariaLabel="착장 날짜 선택" cursor={cursor} onCursorChange={setCursor} selectedDate={selectedDate} onSelectDate={(date) => { setSelectedDate(date); setCalendarOpen(false); }} outfits={outfits}/>}
      </aside>
      <button className={`calendar-fab${calendarOpen ? " is-open" : ""}`} type="button" onClick={() => setCalendarOpen((current) => !current)} onKeyDown={(event) => { if (event.key === "Escape") setCalendarOpen(false); }} aria-label={calendarOpen ? "달력 닫기" : "달력 열기"} aria-expanded={calendarOpen} aria-controls="outfit-calendar-popover" title={calendarOpen ? "달력 닫기" : "달력 열기"}>{calendarOpen ? <X size={22} weight="bold"/> : <CalendarDots size={24}/>}</button>
    </main>}

    {view === "closet" && <main className="collection-page">
      <header className="collection-header"><div><h1>내 옷장</h1><p>{items.length ? "등록한 옷을 관리하고 정보를 수정할 수 있어요." : "자주 입는 옷부터 하나씩 나만의 옷장을 만들어보세요."}</p></div>{items.length > 0 && <div className="header-actions"><button className="secondary-button compact" onClick={openWardrobeDataSettings}><FileXls size={18}/>Excel 관리</button><button className="primary-button compact" onClick={openNewItem}><Plus size={18}/>새 옷 등록</button></div>}</header>
      {items.length ? <>
        <div className="closet-view-toolbar">
          <div className="category-tabs large" aria-label="옷장 카테고리">{allCategories.map((name) => <button type="button" className={category === name ? "active" : ""} aria-pressed={category === name} onClick={() => setCategory(name)} key={name}>{name}</button>)}</div>
          <div className="closet-toolbar-actions"><SearchField value={wardrobeSearch} onChange={setWardrobeSearch} label="옷장 검색" placeholder="이름, 브랜드, 색상 검색"/><SortControl label="옷장 정렬" value={wardrobeSort} onChange={setWardrobeSort}><option value="recent">최근 등록순</option><option value="oldest">오래된 등록순</option><option value="name">이름순</option><option value="category">카테고리순</option><option value="brand">브랜드순</option></SortControl><div className="view-switcher" role="group" aria-label="옷장 보기 방식"><button type="button" className={wardrobeView === "grid" ? "active" : ""} aria-label="그리드로 보기" aria-pressed={wardrobeView === "grid"} onClick={() => setWardrobeView("grid")} title="그리드로 보기"><SquaresFour size={18}/></button><button type="button" className={wardrobeView === "table" ? "active" : ""} aria-label="테이블로 보기" aria-pressed={wardrobeView === "table"} onClick={() => setWardrobeView("table")} title="테이블로 보기"><ListBullets size={19}/></button></div></div>
        </div>
        {sortedFilteredItems.length ? (wardrobeView === "grid"
          ? <section ref={wardrobeCollectionRef} className="collection-grid" aria-label="옷장 그리드">{sortedFilteredItems.map((item) => <button className="collection-card" key={item.id} onClick={() => openEditItem(item.id)}><div className="collection-visual"><ItemVisual item={item}/></div><div><strong>{item.name}</strong><p>{item.brand || "브랜드 미지정"}</p><small>{item.category} · {item.color || "색상 미지정"}</small></div><span className="edit-hint"><PencilSimple size={15}/>수정</span></button>)}</section>
          : <section ref={wardrobeCollectionRef} className="wardrobe-table-shell" aria-label="옷장 테이블"><table className="wardrobe-table"><thead><tr><th scope="col">옷</th><th scope="col">분류</th><th scope="col">브랜드</th><th scope="col">색상</th><th scope="col">상품 링크</th><th scope="col"><span className="sr-only">관리</span></th></tr></thead><tbody>{sortedFilteredItems.map((item) => <tr key={item.id}><td className="wardrobe-primary-cell" data-label="옷"><div className="wardrobe-table-visual"><ItemVisual item={item}/></div><div><strong>{item.name}</strong><small>{item.url ? "상품 정보 연결됨" : "직접 등록한 옷"}</small></div></td><td data-label="분류"><span className="wardrobe-category-chip">{item.category}</span></td><td data-label="브랜드">{item.brand || <span className="table-empty-value">미지정</span>}</td><td data-label="색상">{item.color || <span className="table-empty-value">미지정</span>}</td><td data-label="상품 링크">{item.url ? <a className="wardrobe-product-link" href={item.url} target="_blank" rel="noreferrer"><LinkIcon size={16}/>상품 보기</a> : <span className="table-empty-value">없음</span>}</td><td className="wardrobe-table-action"><button type="button" className="icon-button subtle" onClick={() => openEditItem(item.id)} aria-label={`${item.name} 수정`} title="옷 정보 수정"><PencilSimple size={17}/></button></td></tr>)}</tbody></table></section>)
          : <SearchEmptyState subject="옷" query={wardrobeSearch} category={category} onClear={() => setWardrobeSearch("")} onShowAll={() => setCategory("전체")}/>}
      </> : <section className="collection-empty-state"><span className="empty-state-icon"><Hanger size={28} weight="light"/></span><div className="empty-state-copy"><h2>내 옷장을 시작해볼까요?</h2><p>상품 URL을 붙여 넣으면 정보를 자동으로 채워 첫 옷을 빠르게 등록할 수 있어요.</p></div><div className="empty-state-actions"><button className="primary-button compact" onClick={openNewItem}><Plus size={18}/>첫 옷 등록</button><button className="secondary-button compact" onClick={openWardrobeDataSettings}><FileXls size={18}/>Excel로 한 번에 등록</button></div></section>}
    </main>}

    {view === "lookbook" && <main className="collection-page lookbook-page">
      <header className="collection-header"><div><h1>룩북</h1><p>내 옷으로 구성한 룩과 온라인에서 발견한 코디 영감을 한곳에 모아보세요.</p></div><div className="header-actions"><button className="secondary-button compact" onClick={openInspirationCreator}><ImageSquare size={18}/>스크랩</button>{lookbooks.length > 0 && <button className="primary-button compact" onClick={() => openLookbookCreator()}><Plus size={19}/>새 룩북</button>}</div></header>
      <div className="lookbook-content-layout">
        <section className="owned-lookbooks" aria-labelledby="owned-lookbooks-title">
          <div className="lookbook-section-heading owned-lookbook-toolbar"><div><h2 id="owned-lookbooks-title">내 옷 조합</h2><p>옷장 속 아이템으로 만든, 날짜에 적용할 수 있는 룩이에요.</p></div>{lookbooks.length > 0 && <div className="lookbook-heading-tools"><SearchField value={lookbookSearch} onChange={setLookbookSearch} label="룩북 검색" placeholder="룩북 이름, 메모, 포함된 옷 검색"/><SortControl label="룩북 정렬" value={lookbookSort} onChange={setLookbookSort}><option value="recent">최근 등록순</option><option value="oldest">오래된 등록순</option><option value="name">이름순</option><option value="items">아이템 많은 순</option></SortControl></div>}</div>
          {lookbooks.length ? <>
            {sortedLookbooks.length ? <div className="lookbook-grid">{sortedLookbooks.map((lookbook) => { const lookbookItems = lookbook.itemIds.map((id) => itemMap[id]).filter(Boolean); return <article className="lookbook-card" key={lookbook.id}><div className={`lookbook-mosaic count-${Math.min(lookbookItems.length, 4)}`}>{lookbookItems.slice(0, 4).map((item) => <div className="lookbook-piece" key={item.id}><ItemVisual item={item}/></div>)}</div><div className="lookbook-card-body"><div className="lookbook-title-row"><h2>{lookbook.name}</h2><button className="icon-button subtle" onClick={() => openLookbookDelete(lookbook)} aria-label={`${lookbook.name} 삭제`} title="룩북 삭제"><Trash size={17}/></button></div><p>{lookbook.memo || "다시 입고 싶은 조합으로 저장한 룩북이에요."}</p><div className="lookbook-item-list">{lookbookItems.map((item) => <span key={item.id}>{item.name}</span>)}</div><footer><span>{lookbookItems.length}개의 아이템</span><div className="lookbook-card-actions"><button className="secondary-button compact" onClick={() => openLookbookEditor(lookbook.id)}><PencilSimple size={18}/>수정</button><button className="primary-button compact" onClick={() => openLookbookApply(lookbook.id)}><CalendarPlus size={18}/>날짜에 적용</button></div></footer></div></article>; })}</div> : <SearchEmptyState subject="룩북" query={lookbookSearch} onClear={() => setLookbookSearch("")}/>} {/* 룩북 검색 결과 */}
          </> : <section className="collection-empty-state lookbook-column-empty"><span className="empty-state-icon"><BookOpen size={28} weight="light"/></span><div className="empty-state-copy"><h2>{items.length ? "첫 룩북을 만들어볼까요?" : "룩북을 만들 옷이 필요해요"}</h2><p>{items.length ? "자주 입고 싶은 조합을 구성해 룩북으로 저장하세요." : "먼저 옷을 등록한 뒤 마음에 드는 조합을 룩북으로 저장해보세요."}</p></div><div className="empty-state-actions"><button className="primary-button compact" onClick={items.length ? () => openLookbookCreator() : openNewItem}><Plus size={18}/>{items.length ? "첫 룩북 만들기" : "첫 옷 등록"}</button></div></section>}
        </section>
        <aside className="inspiration-board" aria-labelledby="inspiration-board-title">
          <div className="lookbook-section-heading inspiration-heading"><div><h2 id="inspiration-board-title">스크랩</h2></div><button className="icon-button subtle" onClick={openInspirationCreator} aria-label="새 스크랩 저장" title="스크랩 추가"><Plus size={18}/></button></div>
          {inspirations.length ? <div className="inspiration-masonry">{inspirations.map((pin) => <article className="inspiration-pin" key={pin.id}><div className="inspiration-pin-visual"><button type="button" className="inspiration-image-button" onClick={() => openInspirationViewer(pin)} aria-label={`${pin.caption || "저장한 코디"} 크게 보기`}><img src={pin.image} alt={pin.caption || "저장한 코디 참고 이미지"}/></button><button className="pin-delete-button" onClick={() => openInspirationDelete(pin)} aria-label={`${pin.caption || "스크랩"} 삭제`} title="스크랩 삭제"><Trash size={16}/></button>{pin.sourceUrl && <a href={pin.sourceUrl} target="_blank" rel="noreferrer" aria-label="원본 링크 열기" title="원본 링크 열기"><LinkIcon size={16}/></a>}</div>{pin.caption && <p>{pin.caption}</p>}</article>)}</div> : <button className="inspiration-empty" onClick={openInspirationCreator}><span><ImageSquare size={25}/></span><strong>첫 코디를 스크랩해보세요</strong><small>이미지를 복사해 붙여넣거나 파일·URL로 저장할 수 있어요.</small></button>}
        </aside>
      </div>
    </main>}

    {view === "wishlist" && <main className="collection-page"><header className="collection-header"><div><h1>위시리스트</h1><p>사고 싶은 아이템의 링크와 이유를 가볍게 모아두세요.</p></div>{wishlist.length > 0 && <button className="primary-button compact" onClick={openNewWish}><Plus size={19}/>아이템 스크랩</button>}</header>{wishlist.length ? <><div className="collection-sort-row has-search wishlist-tools"><span>아이템 {wishlist.length}개</span><SearchField value={wishlistSearch} onChange={setWishlistSearch} label="위시리스트 검색" placeholder="이름, 브랜드, 분류, 색상 검색"/><SortControl label="위시리스트 정렬" value={wishlistSort} onChange={setWishlistSort}><option value="recent">최근 등록순</option><option value="oldest">오래된 등록순</option><option value="name">이름순</option><option value="brand">브랜드순</option></SortControl></div>{sortedWishlist.length ? <section className="wishlist-list">{sortedWishlist.map((item) => <article className="wish-row" key={item.id}><div className="wish-visual"><ItemVisual item={item}/></div><div className="wish-copy"><span className="status-chip">{item.status}</span><h3>{item.name}</h3><p>{item.brand || "브랜드 미지정"} · {item.category}{item.color ? ` · ${item.color}` : ""}</p></div>{item.url && <a href={item.url} target="_blank" rel="noreferrer"><LinkIcon size={18}/>상품 보기</a>}<button className="icon-button subtle" onClick={() => openWishlistDelete(item)} aria-label={item.name + " 삭제"} title="위시리스트 삭제"><Trash size={18}/></button></article>)}</section> : <SearchEmptyState subject="아이템" query={wishlistSearch} onClear={() => setWishlistSearch("")}/>}</> : <section className="collection-empty-state wishlist-empty-state"><span className="empty-state-icon"><Heart size={28} weight="light"/></span><div className="empty-state-copy"><h2>마음에 둔 상품이 있나요?</h2><p>상품 URL을 저장하면 이름과 이미지를 자동으로 가져와 나중에 다시 볼 수 있어요.</p></div><div className="empty-state-actions"><button className="primary-button compact" onClick={openNewWish}><Plus size={18}/>첫 아이템 저장</button></div></section>}</main>}

    {view === "settings" && <main className="settings-page"><header className="settings-header"><h1>설정</h1><p>화면 분위기와 옷장 데이터 구조를 한곳에서 관리하세요.</p></header><div className="settings-layout"><nav className="settings-tabs"><button className={settingsTab === "general" ? "active" : ""} onClick={() => setSettingsTab("general")}>일반</button><button className={settingsTab === "closet" || settingsTab === "excel" ? "active" : ""} onClick={() => setSettingsTab("closet")}>옷장</button><button className={settingsTab === "backup" ? "active" : ""} onClick={() => setSettingsTab("backup")}>백업</button><button className={settingsTab === "reset" ? "active" : ""} onClick={() => setSettingsTab("reset")}>초기화</button></nav><section className="settings-content">
      {settingsTab === "general" && <div className="settings-panel"><div className="panel-heading"><div><h2>화면 모드</h2><p>필요할 때 언제든 라이트와 다크 모드를 바꿀 수 있어요.</p></div></div><div className="theme-options"><button className={theme === "light" ? "selected" : ""} onClick={() => setTheme("light")}><Sun size={24}/><span><strong>라이트</strong><small>밝고 여백감 있는 화면</small></span>{theme === "light" && <CheckCircle size={21} weight="fill"/>}</button><button className={theme === "dark" ? "selected" : ""} onClick={() => setTheme("dark")}><Moon size={24}/><span><strong>다크</strong><small>차분하고 눈이 편한 화면</small></span>{theme === "dark" && <CheckCircle size={21} weight="fill"/>}</button></div></div>}
      {settingsTab === "closet" && <div className="settings-panel"><div className="panel-heading"><div><h2>카테고리 관리</h2><p>한 줄 목록에서 드래그하거나 위아래 화살표로 순서를 바꿔보세요. 변경한 순서는 착장과 옷장에 바로 반영됩니다.</p></div></div><form className="inline-form" onSubmit={addCategory}><input name="categoryName" aria-label="새 카테고리 이름" placeholder="예: 운동복"/><button className="primary-button compact"><Plus size={18}/>추가</button></form><div className="category-manager">{categories.map((name, index) => { const used = items.some((item) => item.category === name); return <div className={`category-row${draggedCategoryName === name ? " is-dragging" : ""}${categoryDropTarget === name && draggedCategoryName !== name ? " is-drop-target" : ""}`} key={name} onDragEnter={() => setCategoryDropTarget(name)} onDragOver={(event) => { event.preventDefault(); event.dataTransfer.dropEffect = "move"; }} onDrop={(event) => finishCategoryDrop(event, name)}><button type="button" className="category-drag-handle" draggable onDragStart={(event) => startCategoryDrag(event, name)} onDragEnd={() => { setDraggedCategoryName(null); setCategoryDropTarget(null); }} onKeyDown={(event) => handleCategoryKeyDown(event, name)} aria-label={`${name} 순서 변경. 위아래 화살표 키 사용`} title="드래그해서 순서 변경"><DotsSixVertical size={18} weight="bold"/></button><span>{name}<small>{items.filter((item) => item.category === name).length}개</small></span><div className="category-order-actions"><button type="button" className="category-order-button" onClick={() => moveCategory(name, -1)} aria-label={`${name} 위로 이동`} disabled={index === 0}><CaretUp size={15} weight="bold"/></button><button type="button" className="category-order-button" onClick={() => moveCategory(name, 1)} aria-label={`${name} 아래로 이동`} disabled={index === categories.length - 1}><CaretDown size={15} weight="bold"/></button><button type="button" className="icon-button subtle" onClick={() => deleteCategory(name)} aria-label={`${name} 삭제`} disabled={used} title={used ? "사용 중인 카테고리" : "카테고리 삭제"}><Trash size={17}/></button></div></div>; })}</div></div>}
      {settingsTab === "closet" && <div className="settings-panel excel-panel" id="wardrobe-data-management" ref={wardrobeDataRef} tabIndex={-1}><div className="panel-heading"><div><h2>옷장 데이터 관리</h2><p>옷장 데이터를 Excel로 가져오거나 현재 데이터를 파일로 저장할 수 있어요.</p></div><FileXls size={32}/></div><div className="excel-actions"><section className="excel-action-card"><div><strong>Excel에서 옷 가져오기</strong><small>최신 템플릿을 작성한 뒤 업로드하면 옷장 목록에 추가돼요.</small></div><div className="excel-action-controls"><a className="secondary-button compact" href={TEMPLATE_URL} download><DownloadSimple size={18}/>템플릿 받기</a><FilePicker id="excel-upload" name="excelFile" accept=".xlsx,.xls" fileName={excelFileName} label="Excel 가져오기" icon="excel" compact tone="primary" onChange={(event) => importExcel(event.target.files?.[0])}/></div></section><section className="excel-action-card"><div><strong>현재 옷장 내보내기</strong><small>{items.length ? `등록된 옷 ${items.length}개의 정보를 Excel 파일로 저장해요.` : "등록된 옷이 없어요. 버튼을 누르면 필요한 작업을 안내해드려요."}</small></div><button type="button" className="secondary-button compact" onClick={exportExcel}><DownloadSimple size={18}/>Excel 내보내기</button></section></div><p className="excel-notice">Excel은 옷장 데이터만 관리합니다. 착장, 메모, 룩북, 위시리스트는 포함되지 않아요. 직접 업로드한 이미지 파일은 내보내지 않으며 웹 이미지 URL만 유지합니다.</p></div>}
      {settingsTab === "backup" && <div className="settings-panel backup-panel"><div className="panel-heading"><div><h2>전체 데이터 백업</h2><p>브라우저를 옮기거나 데이터를 안전하게 보관할 때 전체 백업을 사용하세요.</p></div><DownloadSimple size={32}/></div><div className="backup-actions"><section className="excel-action-card"><div><strong>전체 데이터 내보내기</strong><small>옷과 이미지, 착장, 메모, 룩북, 스크랩, 위시리스트 및 화면 설정을 JSON 파일로 저장해요.</small></div><button type="button" className="secondary-button compact" onClick={exportBackup}><DownloadSimple size={18}/>백업 다운로드</button></section><section className="excel-action-card"><div><strong>백업 파일 가져오기</strong><small>선택한 백업의 내용으로 현재 브라우저 데이터를 교체해요.</small></div><FilePicker id="backup-upload" name="backupFile" accept=".json,application/json" fileName={backupFileName} label="백업 가져오기" compact tone="primary" onChange={(event) => prepareBackupImport(event.target.files?.[0])}/></section></div><p className="backup-notice">업로드한 이미지와 스크랩도 백업에 포함됩니다. 가져오기는 현재 데이터를 교체하므로 필요한 경우 먼저 백업을 내려받아 안전하게 보관하세요.</p></div>}
      {settingsTab === "reset" && <div className="settings-panel settings-danger-zone"><div className="panel-heading"><div><h2>데이터 초기화</h2><p>옷장을 처음부터 다시 정리합니다. 옷을 기준으로 저장된 착장과 메모, 룩북도 함께 삭제됩니다.</p></div></div><div className="danger-zone-action"><div><strong>옷장 데이터 초기화</strong><small>위시리스트와 화면 모드·사이드바 설정은 유지됩니다.</small></div><button type="button" className="danger-button" onClick={() => setModal("confirmResetWardrobe")}><Trash size={18}/>초기화</button></div></div>}
    </section></div></main>}

    {modal === "outfit" && <Modal title="착장 고르기" subtitle={selectedDateLabel} eyebrow="" onClose={closeModal} wide><div className="category-tabs modal-tabs">{allCategories.map((name) => <button className={outfitCategory === name ? "active" : ""} onClick={() => setOutfitCategory(name)} key={name}>{name}</button>)}</div><div className="outfit-picker-body">{outfitPickerItems.length ? <div className="picker-grid">{outfitPickerItems.map((item) => { const checked = draftOutfit.includes(item.id); return <button type="button" className={checked ? "selected" : ""} onClick={() => toggleDraft(item.id)} key={item.id}><div className="picker-visual"><ItemVisual item={item}/></div><span><strong>{item.name}</strong><small>{item.brand || "브랜드 미지정"}</small><small>{item.category} · {item.color || "색상 미지정"}</small></span>{checked && <i className="check"><Check size={15} weight="bold"/></i>}</button>; })}</div> : <div className="picker-empty">이 카테고리에는 등록된 옷이 없어요.</div>}</div><footer className="modal-actions">{selectedItems.length > 0 && <button className="danger-button outfit-modal-delete" onClick={() => setModal("confirmOutfit")}><Trash size={18}/>착장 삭제</button>}<span className="action-spacer"/><button className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact" onClick={saveOutfit}>착장 저장</button></footer></Modal>}
    {modal === "lookbookCreate" && <Modal title={editingLookbookId ? "룩북 수정" : "룩북 만들기"} subtitle={editingLookbookId ? "이름과 메모, 옷 구성을 변경하세요" : "다시 입고 싶은 조합을 한곳에 저장해두세요"} eyebrow="" onClose={closeModal} wide className="lookbook-editor-modal"><div className="category-tabs modal-tabs" aria-label="옷 카테고리">{allCategories.map((name) => <button type="button" className={lookbookCategory === name ? "active" : ""} aria-pressed={lookbookCategory === name} onClick={() => setLookbookCategory(name)} key={name}>{name}</button>)}</div><div className="lookbook-picker-summary" aria-live="polite"><strong>{draftLookbookIds.length}개 선택</strong><span>함께 입을 옷을 골라 하나의 룩으로 구성해보세요.</span></div><div className="outfit-picker-body">{lookbookPickerItems.length ? <div className="picker-grid">{lookbookPickerItems.map((item) => { const checked = draftLookbookIds.includes(item.id); return <button type="button" className={checked ? "selected" : ""} aria-pressed={checked} onClick={() => toggleLookbookDraft(item.id)} key={item.id}><div className="picker-visual"><ItemVisual item={item}/></div><span><strong>{item.name}</strong><small>{item.brand || "브랜드 미지정"}</small><small>{item.category} · {item.color || "색상 미지정"}</small></span>{checked && <i className="check"><Check size={15} weight="bold"/></i>}</button>; })}</div> : <div className="picker-empty">이 카테고리에는 등록된 옷이 없어요.</div>}</div><div className="lookbook-form-fields"><label><FieldTitle required>룩북 이름</FieldTitle><input value={lookbookName} onChange={(event) => setLookbookName(event.target.value)} placeholder="예: 비 오는 날 출근룩"/></label><label><FieldTitle>메모</FieldTitle><input value={lookbookMemo} onChange={(event) => setLookbookMemo(event.target.value)} placeholder="언제, 어떤 분위기로 입을지 적어두세요"/></label></div><footer className="modal-actions"><button className="secondary-button compact" onClick={closeModal}>취소</button><button className="primary-button compact" onClick={saveLookbook}>{editingLookbookId ? "수정 저장" : "룩북 저장"}</button></footer></Modal>}
    {modal === "inspirationCreate" && <Modal title="스크랩 저장" subtitle="온라인에서 발견한 스타일을 참고용으로 모아두세요" eyebrow="" onClose={closeModal} compact className="inspiration-modal"><form className="entry-form inspiration-form" onSubmit={saveInspiration}><div className={`pin-paste-zone${pinImage ? " has-image" : ""}`} tabIndex={0}>{pinImage ? <><img src={pinImage} alt="저장할 코디 미리보기"/><button type="button" className="pin-preview-remove" onClick={() => { setPinImage(""); setPinFileName(""); }} aria-label="선택한 이미지 제거"><X size={16}/></button></> : <><span><ImageSquare size={28}/></span><strong>이미지를 여기에 붙여넣으세요</strong><small>스크린샷을 복사한 뒤 ⌘V 또는 Ctrl+V</small></>}</div><div className="field-group"><FieldTitle required>이미지 URL 또는 파일</FieldTitle><div className="image-source-row"><div className="input-with-icon"><ImageSquare size={19}/><input name="imageUrl" type="url" inputMode="url" disabled={Boolean(pinImage)} placeholder="https://image.example.com/look.jpg"/></div><input id="pin-image-file" className="visually-hidden-input" type="file" accept="image/*" onChange={(event) => handlePinFile(event.target.files?.[0])}/><label className="image-upload-button" htmlFor="pin-image-file"><UploadSimple size={18}/><span>파일</span></label></div><small className={`field-help${pinFileName ? " selected" : ""}`}>{pinFileName ? `${pinFileName} 선택됨 · 저장할 때 자동으로 용량을 줄여요.` : "JPG, PNG, WEBP · 최대 2.5MB · 저장할 때 자동으로 용량을 줄여요."}</small></div><label><FieldTitle>캡션</FieldTitle><input value={pinCaption} onChange={(event) => setPinCaption(event.target.value)} maxLength={80} placeholder="예: 여름 여행용 뉴트럴 레이어드"/></label><label><FieldTitle>원본 URL</FieldTitle><div className="input-with-icon"><LinkIcon size={19}/><input type="url" value={pinSourceUrl} onChange={(event) => setPinSourceUrl(event.target.value)} placeholder="https://pinterest.com/…"/></div></label><p className="pin-reference-note">스크랩은 참고용으로만 저장되며 착장 날짜에는 적용되지 않아요.</p><footer className="modal-actions"><button type="button" className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact">스크랩 저장</button></footer></form></Modal>}
    {modal === "inspirationViewer" && activeInspiration && <InspirationViewer inspiration={activeInspiration} onClose={closeModal}/>}
    {modal === "lookbookApply" && activeLookbook && <Modal title="날짜에 적용" subtitle={activeLookbook.name} eyebrow="" onClose={closeModal} compact className="lookbook-apply-modal"><div className="lookbook-apply-copy"><CalendarPlus size={28}/><div><strong>{activeLookbook.itemIds.filter((id) => itemMap[id]).length}개의 아이템을 적용할 날짜</strong><span>선택한 날짜의 기존 착장이 있다면 이 룩북으로 교체됩니다.</span></div></div><CompactCalendar className="lookbook-apply-calendar" ariaLabel="룩북을 적용할 날짜 선택" cursor={lookbookCalendarCursor} onCursorChange={setLookbookCalendarCursor} selectedDate={lookbookApplyDate} onSelectDate={setLookbookApplyDate} outfits={outfits}/>{(outfits[lookbookApplyDate] ?? []).length > 0 && <p className="lookbook-overwrite-note">이 날짜에는 이미 착장이 있어요. 적용하면 현재 조합이 교체됩니다.</p>}<footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact" onClick={applyLookbook}>이 날짜에 적용</button></footer></Modal>}
    {modal === "itemUrl" && <Modal title="상품 URL로 옷 등록" subtitle="링크를 분석한 뒤 입력 내용을 확인할 수 있어요" onClose={closeModal} compact className="product-url-modal"><ProductUrlStep kind="item" loading={productLookupLoading} error={productLookupError} onSubmit={(event) => lookupProduct(event, "item")} onManual={() => openManualRegistration("item")}/></Modal>}
    {modal === "wishUrl" && <Modal title="상품 URL로 위시리스트 등록" subtitle="링크를 분석한 뒤 입력 내용을 확인할 수 있어요" onClose={closeModal} compact className="product-url-modal"><ProductUrlStep kind="wish" loading={productLookupLoading} error={productLookupError} onSubmit={(event) => lookupProduct(event, "wish")} onManual={() => openManualRegistration("wish")}/></Modal>}
    {modal === "item" && <Modal title="새 옷 등록" subtitle={productDraft?.autoFilled ? "자동 입력된 내용을 확인해주세요" : "직접 상품 정보를 입력해주세요"} onClose={closeModal} className="item-editor-modal">{renderItemForm({ mode: "create" })}</Modal>}
    {modal === "editItem" && editingItem && <Modal title="옷 정보 수정" onClose={closeModal} className="item-editor-modal">{renderItemForm({ mode: "edit" })}</Modal>}
    {modal === "confirmOutfit" && <Modal title="이 착장을 삭제할까요?" onClose={() => setModal("outfit")} compact><div className="confirm-copy"><Trash size={24}/><p><strong>{selectedDateLabel} 착장과 메모가 삭제됩니다.</strong><span>삭제 후에는 되돌릴 수 없어요.</span></p></div><footer className="modal-actions"><button className="secondary-button" onClick={() => setModal("outfit")}>취소</button><button className="danger-button filled" onClick={deleteOutfit}>착장 삭제</button></footer></Modal>}
    {modal === "confirmItem" && editingItem && <Modal title="이 옷을 삭제할까요?" onClose={() => setModal("editItem")} compact><div className="confirm-copy"><Trash size={24}/><p><strong>{editingItem.name}이(가) 옷장에서 삭제됩니다.</strong><span>저장된 착장에서도 함께 제거되며 되돌릴 수 없어요.</span></p></div><footer className="modal-actions"><button className="secondary-button" onClick={() => setModal("editItem")}>취소</button><button className="danger-button filled" onClick={deleteItem}>옷 삭제</button></footer></Modal>}
    {modal === "confirmLookbook" && pendingDeletion?.type === "lookbook" && <Modal title="이 룩북을 삭제할까요?" onClose={closeModal} compact><div className="confirm-copy"><Trash size={24}/><p><strong>{pendingDeletion.name} 룩북이 삭제됩니다.</strong><span>이미 날짜에 적용한 착장은 유지되며, 룩북 삭제 후에는 되돌릴 수 없어요.</span></p></div><footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="danger-button filled" onClick={deleteLookbook}>룩북 삭제</button></footer></Modal>}
    {modal === "confirmWishlist" && pendingDeletion?.type === "wishlist" && <Modal title="이 아이템을 삭제할까요?" onClose={closeModal} compact><div className="confirm-copy"><Trash size={24}/><p><strong>{pendingDeletion.name}이(가) 위시리스트에서 삭제됩니다.</strong><span>삭제 후에는 되돌릴 수 없어요.</span></p></div><footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="danger-button filled" onClick={deleteWishlistItem}>아이템 삭제</button></footer></Modal>}
    {modal === "confirmInspiration" && pendingDeletion?.type === "inspiration" && <Modal title="이 스크랩을 삭제할까요?" onClose={closeModal} compact><div className="confirm-copy"><Trash size={24}/><p><strong>{pendingDeletion.name} 스크랩이 삭제됩니다.</strong><span>삭제 후에는 되돌릴 수 없어요.</span></p></div><footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="danger-button filled" onClick={deleteInspiration}>스크랩 삭제</button></footer></Modal>}
    {modal === "confirmResetWardrobe" && <Modal title="옷장 데이터를 초기화할까요?" onClose={closeModal} compact><div className="confirm-copy reset-confirm-copy"><Trash size={24}/><div><strong>연결된 데이터가 모두 삭제됩니다.</strong><ul><li>옷 {items.length}개와 커스텀 카테고리</li><li>착장 {Object.values(outfits).filter((ids) => ids.length).length}일과 날짜별 메모</li><li>룩북 {lookbooks.length}개</li></ul><span>위시리스트와 화면 설정은 유지되며, 초기화 후에는 되돌릴 수 없습니다.</span></div></div><footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="danger-button filled" onClick={resetWardrobeData}>모두 초기화</button></footer></Modal>}
    {modal === "confirmBackupImport" && pendingBackup && <Modal title="백업 데이터를 가져올까요?" onClose={closeModal} compact><div className="confirm-copy backup-confirm-copy"><UploadSimple size={24}/><div><strong>{backupFileName}</strong><ul><li>옷 {pendingBackup.data.items.length}개 · 카테고리 {pendingBackup.data.categories.length}개</li><li>착장 {Object.values(pendingBackup.data.outfits).filter((ids) => Array.isArray(ids) && ids.length).length}일 · 룩북 {pendingBackup.data.lookbooks.length}개</li><li>위시리스트 {pendingBackup.data.wishlist.length}개 · 스크랩 {(pendingBackup.data.inspirations ?? []).length}개</li></ul><span>현재 브라우저의 전체 데이터가 이 파일 내용으로 교체됩니다. 필요하다면 먼저 현재 데이터를 내보내 주세요.</span></div></div><footer className="modal-actions"><button className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact" onClick={importBackup}>가져오기</button></footer></Modal>}
    {modal === "wish" && <Modal title="아이템 스크랩" subtitle={productDraft?.autoFilled ? "자동 입력된 내용을 확인해주세요" : "직접 상품 정보를 입력해주세요"} onClose={closeModal} className="wish-editor-modal"><AutofillSummary product={productDraft}/><form className="entry-form" onSubmit={addWish}><label><FieldTitle required>아이템 이름</FieldTitle><input name="name" required defaultValue={productDraft?.name || ""} placeholder="예: 브라운 레더 토트백"/></label><div className="form-row"><label><FieldTitle>브랜드</FieldTitle><input name="brand" defaultValue={productDraft?.brand || ""} placeholder="예: Aesther Ekme"/></label><label><FieldTitle required>분류</FieldTitle><SelectField name="category" required defaultValue={productDraft?.category || categories[0]}>{categories.map((name) => <option key={name}>{name}</option>)}</SelectField></label></div><label><FieldTitle>색상</FieldTitle><input name="color" defaultValue={productDraft?.color || ""} placeholder="예: 옐로우 다크 블루"/></label><label><FieldTitle>상품 URL</FieldTitle><div className="input-with-icon"><LinkIcon size={19}/><input name="url" type="url" defaultValue={productDraft?.url || ""} placeholder="https://"/></div></label><div className="field-group image-source-field"><FieldTitle>이미지 URL 또는 파일</FieldTitle><div className="image-source-row"><div className="input-with-icon"><ImageSquare size={19}/><input name="image" type="text" inputMode="url" defaultValue={productDraft?.image || ""} placeholder="https://image.example.com/item.jpg"/></div><input id="wish-image" className="visually-hidden-input" name="imageFile" type="file" accept="image/*" onChange={(event) => setWishFileName(event.target.files?.[0]?.name || "")}/><label className="image-upload-button" htmlFor="wish-image" title="이미지 파일 선택"><UploadSimple size={18}/><span>파일</span></label></div><small className={`field-help image-source-help${wishFileName ? " selected" : ""}`} aria-live="polite">{wishFileName ? `${wishFileName} 선택됨 · URL 이미지 대신 이 파일을 사용해요.` : "이미지 URL을 입력하거나 JPG, PNG, WEBP 파일을 선택하세요. 최대 2.5MB"}</small></div><footer className="modal-actions"><button type="button" className="secondary-button" onClick={closeModal}>취소</button><button className="primary-button compact">스크랩 저장</button></footer></form></Modal>}
    {dragPreview && itemMap[dragPreview.id] && <div className="drag-preview" style={{ "--drag-x": `${dragPreview.x}px`, "--drag-y": `${dragPreview.y}px` }} aria-hidden="true"><ItemVisual item={itemMap[dragPreview.id]}/><strong>{itemMap[dragPreview.id].name}</strong></div>}
    {notice && <div className="toast" role="status"><CheckCircle size={20} weight="fill"/>{notice}</div>}
  </div>;
}
