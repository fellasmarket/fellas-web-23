import {
  useState,
  useMemo,
  useEffect,
  useRef,
  FormEvent,
  ChangeEvent,
  DragEvent,
} from "react";

import { useQueryClient } from "@tanstack/react-query";
import {
  useGetMenu,
  useUpdateSettings,
  useCreateProduct,
  useUpdateProduct,
  useDeleteProduct,
  useCreateCategory,
  useDeleteCategory,
  useReorderCategories,
  useCreateAisle,
  useDeleteAisle,
  useUpdateAisle,
  useCreateSubcategory,
  useDeleteSubcategory,
  useAdminLogin,
  useListOrders,
  useCreateOrder,
  useDeleteOrder,
  useUpdateOrderStatus,
  useTrackVisit,
  useGetVisitStats,
  useGetSalesStats,
  useReorderAisles,
  useReorderSubcategories,
  useReorderProducts,
  getGetMenuQueryKey,
  getListOrdersQueryKey,
  Product,
  Settings,
  Category,
  Order,
  BannerSlide,
  DeliveryLocation,
  useListDeliveryLocations,
  useCreateDeliveryLocation,
  useUpdateDeliveryLocation,
  useDeleteDeliveryLocation,
  getListDeliveryLocationsQueryKey,
} from "@workspace/api-client-react";
import { generateStatsPDF } from "../lib/statsPdf";
import {
  ShoppingCart,
  Settings as SettingsIcon,
  Plus,
  Trash2,
  Phone,
  LogOut,
  X,
  Utensils,
  Edit3,
  Upload,
  Search,
  GripVertical,
  CheckCircle,
  Layers,
  Tag,
  MapPin,
  ClipboardList,
  User,
  BarChart3,
  TrendingUp,
  Globe,
  Clock,
  Calendar,
  Eye,
  EyeOff,
  Flame,
  ChevronDown,
  ChevronRight,
  Download,
  Instagram,
  ArrowUp,
  ArrowDown,
  ChevronUp,
  Minus,
  Sparkles,
  MapPinIcon,
  PhoneCall,
  ClockIcon,
  Link2,
  HardDriveUpload,
  Ban,
  Pencil,
  CreditCard,
  Share2,
  ImageIcon,
  Copy,
  UserPlus,
  Percent,
  Bookmark,
  KeyRound,
  Gift,
  Users,
  Loader2,
  ShoppingBag,
} from "lucide-react";
import * as XLSX from "xlsx";

function downloadProductsExcel(products: Product[]) {
  const sorted = [...products].sort((a, b) => {
    const cat = (a.category ?? "").localeCompare(b.category ?? "", "es");
    if (cat !== 0) return cat;
    const aisle = (a.aisle ?? "").localeCompare(b.aisle ?? "", "es");
    if (aisle !== 0) return aisle;
    const sub = (a.subcategory ?? "").localeCompare(b.subcategory ?? "", "es");
    if (sub !== 0) return sub;
    return (a.name ?? "").localeCompare(b.name ?? "", "es");
  });

  const headers = ["codigo", "nombre", "variante", "categoria", "proveedor", "precio", "costo", "stock", "stock_minimo"];
  const rows = sorted.map(p => [
    p.id ?? "",
    p.name ?? "",
    p.subcategory ?? "",
    p.category ?? "",
    p.aisle ?? "",
    p.price ?? 0,
    "",
    "",
    "",
  ]);

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  ws["!cols"] = [
    { wch: 10 }, { wch: 30 }, { wch: 25 }, { wch: 18 }, { wch: 20 },
    { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 12 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Productos");
  XLSX.writeFile(wb, "plantilla_productos_POS.xlsx");
}

const subcategoryEmoji = (name: string): string => {
  const n = name.toLowerCase();
  if (/cerveza|chela|beer|lager|ipa|stout/.test(n)) return "🍺";
  if (/champ|espumant|sparkl|prosec/.test(n)) return "🍾";
  if (/vino|tinto|blanco|rosado|cabernet|merlot|carmen/.test(n)) return "🍷";
  if (/pisco|whisky|whiskey|ron|vodka|tequila|gin|destil|licor/.test(n)) return "🥃";
  if (/cocktail|coctel|trago|mix|piscola|fernet/.test(n)) return "🍹";
  if (/sake|asia/.test(n)) return "🍶";
  if (/snack|papa|chips|frito|maní|mani|nuez|fruto/.test(n)) return "🍿";
  if (/dulce|chocolate|caramelo|gomit|galleta/.test(n)) return "🍫";
  if (/agua|jugo|gaseos|bebid|soda|cola/.test(n)) return "🥤";
  if (/energ|red bull|monster/.test(n)) return "⚡";
  if (/cigarr|tabac|vape/.test(n)) return "🚬";
  if (/hielo|cold|frio/.test(n)) return "🧊";
  if (/promo|oferta|combo/.test(n)) return "🎉";
  if (/import/.test(n)) return "🌎";
  if (/nacional|chile/.test(n)) return "🇨🇱";
  return "✨";
};

const normalize = (s: string) =>
  s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

function levenshtein(a: string, b: string): number {
  if (Math.abs(a.length - b.length) > 3) return 99;
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const prev = Array.from({ length: n + 1 }, (_, i) => i);
  const curr = new Array<number>(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      curr[j] = a[i - 1] === b[j - 1]
        ? prev[j - 1]
        : 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
    }
    for (let j = 0; j <= n; j++) prev[j] = curr[j];
  }
  return curr[n];
}

function fuzzyMatch(query: string, target: string): boolean {
  if (!query) return true;
  const q = normalize(query);
  const t = normalize(target);
  // Exact substring match first
  if (t.includes(q)) return true;
  // Each query word must match some target word
  const qWords = q.split(/\s+/).filter(Boolean);
  const tWords = t.split(/\s+/).filter(Boolean);
  return qWords.every((qw) =>
    tWords.some((tw) => {
      // Target word starts with what the user typed (prefix match)
      if (tw.startsWith(qw)) return true;
      // Only allow Levenshtein for words of 4+ chars, and be strict
      if (qw.length < 4 || tw.length < 4) return false;
      const maxDist = qw.length >= 7 ? 2 : 1;
      return levenshtein(qw, tw) <= maxDist;
    }),
  );
}

interface CartItem extends Product {
  quantity: number;
  selectedOption: string | null;
  cartItemId: number;
}

const DEFAULT_FORM = {
  name: "",
  price: "",
  image: "",
  category: "",
  aisle: "",
  subcategory: "",
  optionsTitle: "",
  optionsString: "",
  oferta: false,
  depositoEnabled: false,
  depositoAmount: "500",
  transferenciaEnabled: false,
  transferenciaAmount: "0",
};

function useDragScroll() {
  const ref = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    let isDown = false;
    let moved = false;
    let startX = 0;
    let startScroll = 0;
    const onMouseDown = (e: MouseEvent) => {
      isDown = true;
      moved = false;
      startX = e.pageX;
      startScroll = el.scrollLeft;
      el.style.cursor = "grabbing";
    };
    const onMouseUp = () => {
      if (!isDown) return;
      isDown = false;
      el.style.cursor = "grab";
    };
    const onMouseMove = (e: MouseEvent) => {
      if (!isDown) return;
      const dx = e.pageX - startX;
      if (Math.abs(dx) > 5) {
        moved = true;
        el.scrollLeft = startScroll - dx;
        e.preventDefault();
      }
    };
    const onClickCapture = (e: MouseEvent) => {
      if (moved) {
        e.stopPropagation();
        e.preventDefault();
        moved = false;
      }
    };
    const onWheel = (e: WheelEvent) => {
      if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
        el.scrollLeft += e.deltaY;
      }
    };
    el.style.cursor = "grab";
    el.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mouseup", onMouseUp);
    window.addEventListener("mousemove", onMouseMove);
    el.addEventListener("click", onClickCapture, true);
    el.addEventListener("wheel", onWheel, { passive: true });
    return () => {
      el.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("mousemove", onMouseMove);
      el.removeEventListener("click", onClickCapture, true);
      el.removeEventListener("wheel", onWheel);
    };
  }, []);
  return ref;
}

export default function Storefront() {
  const queryClient = useQueryClient();
  const { data: menu, isLoading } = useGetMenu();
  const { data: deliveryLocations = [] } = useListDeliveryLocations();

  const [ageVerified, setAgeVerified] = useState<boolean>(() => {
    try { return sessionStorage.getItem("fella_age_ok") === "1"; } catch { return false; }
  });
  const handleAgeYes = () => {
    try { sessionStorage.setItem("fella_age_ok", "1"); } catch {}
    setAgeVerified(true);
  };
  const handleAgeNo = () => {
    window.location.replace("https://www.google.com");
  };

  const [showScrollTop, setShowScrollTop] = useState(false);
  useEffect(() => {
    const onScroll = () => setShowScrollTop(window.scrollY > 400);
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  const [keyboardOffset, setKeyboardOffset] = useState(0);
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    const update = () => {
      const offset = window.innerHeight - (vv.offsetTop + vv.height);
      setKeyboardOffset(Math.max(0, offset));
    };
    vv.addEventListener("resize", update);
    vv.addEventListener("scroll", update);
    return () => {
      vv.removeEventListener("resize", update);
      vv.removeEventListener("scroll", update);
    };
  }, []);
  const scrollToTop = () => window.scrollTo({ top: 0, behavior: "smooth" });

  const settings: Settings = menu?.settings ?? {
    logo: "",
    pageTitle: "Urban Bite",
    favicon: "",
    bannerText: "",
    bannerDescription: "",
    bannerImage: "",
    bannerSlides: [],
    whatsapp: "",
    agencyName: "",
    agencyLogo: "",
    announcements: [],
    contactPhone: "",
    contactAddress: "",
    contactHours: "",
    openTime: "11:00",
    closeTime: "23:00",
    deliveryMinimum: 10000,
  };
  const categoriesData: Category[] = menu?.categories ?? [];
  const aislesData = menu?.aisles ?? [];
  const subcategoriesData = menu?.subcategories ?? [];
  const products: Product[] = menu?.products ?? [];

  const categories = categoriesData.map((c) => c.name);
  const aisles = aislesData.map((a) => a.name);
  const subcategories = subcategoriesData.map((s) => s.name);

  const updateSettingsMut = useUpdateSettings();
  const createProductMut = useCreateProduct();
  const updateProductMut = useUpdateProduct();
  const deleteProductMut = useDeleteProduct();
  const createCategoryMut = useCreateCategory();
  const deleteCategoryMut = useDeleteCategory();
  const reorderCategoriesMut = useReorderCategories();
  const createAisleMut = useCreateAisle();
  const deleteAisleMut = useDeleteAisle();
  const updateAisleMut = useUpdateAisle();
  const reorderAislesMut = useReorderAisles();
  const [aisleBannerDrafts, setAisleBannerDrafts] = useState<Record<string, string>>({});
  const createSubcategoryMut = useCreateSubcategory();
  const deleteSubcategoryMut = useDeleteSubcategory();
  const reorderSubcategoriesMut = useReorderSubcategories();
  const reorderProductsMut = useReorderProducts();
  const adminLoginMut = useAdminLogin();

  const refreshMenu = () =>
    queryClient.invalidateQueries({ queryKey: getGetMenuQueryKey() });

  const moveProduct = (productId: number, direction: "up" | "down", groupIds: number[]) => {
    const idx = groupIds.indexOf(productId);
    if (idx === -1) return;
    const newIdx = direction === "up" ? idx - 1 : idx + 1;
    if (newIdx < 0 || newIdx >= groupIds.length) return;
    const newGroup = [...groupIds];
    [newGroup[idx], newGroup[newIdx]] = [newGroup[newIdx], newGroup[idx]];
    const groupSet = new Set(groupIds);
    const allIds = products.map((p) => p.id);
    const newAllIds: number[] = [];
    let inserted = false;
    for (const id of allIds) {
      if (!groupSet.has(id)) {
        newAllIds.push(id);
      } else if (!inserted) {
        newAllIds.push(...newGroup);
        inserted = true;
      }
    }
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? {
        ...old,
        products: newAllIds.map((id) => old.products.find((p: any) => p.id === id)).filter(Boolean),
      } : old,
    );
    reorderProductsMut.mutate({ data: { ids: newAllIds } });
  };

  const [view, setView] = useState<"client" | "admin-login" | "admin">("client");
  const [cart, setCart] = useState<CartItem[]>([]);
  const [isCartOpen, setIsCartOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchQuery), 180);
    return () => clearTimeout(t);
  }, [searchQuery]);
  const [toast, setToast] = useState({
    show: false,
    message: "",
    actionLabel: "",
    onAction: null as null | (() => void),
  });
  const [activeCategory, setActiveCategory] = useState("");
  const [activeAisle, setActiveAisle] = useState("");
  const [showAisleMenu, setShowAisleMenu] = useState(false);
  const [checkoutOpen, setCheckoutOpen] = useState(false);
  const [checkoutStep, setCheckoutStep] = useState(1);
  const [checkoutForm, setCheckoutForm] = useState({
    customerName: "",
    address: "",
    phone: "",
    notes: "",
    deliveryType: "retiro" as "delivery" | "retiro",
  });

  const [optionModalInfo, setOptionModalInfo] = useState<
    { product: Product; selectedOption: string } | null
  >(null);
  const [adminTab, setAdminTab] =
    useState<"products" | "classifications" | "orders" | "stats" | "settings" | "social" | "customers">("products");
  const [adminRole, setAdminRole] = useState<"full" | "delivery" | null>(null);
  const [customerToken, setCustomerToken] = useState<string | null>(() => {
    try { return localStorage.getItem("customer_token"); } catch { return null; }
  });
  const [customer, setCustomer] = useState<{ id: number; email: string; name: string; phone: string } | null>(null);
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [authMode, setAuthMode] = useState<"login" | "register">("login");
  const [showAccountPanel, setShowAccountPanel] = useState(false);
  const [savedCarts, setSavedCarts] = useState<Array<{ id: number; name: string; items: any[]; createdAt: string }>>([]);
  const [selectedLocation, setSelectedLocation] = useState<DeliveryLocation | null>(null);
  const [discountCode, setDiscountCode] = useState("");
  const [appliedDiscount, setAppliedDiscount] = useState<{ code: string; type: string; amount: number; discountAmount: number; label: string } | null>(null);
  const [discountError, setDiscountError] = useState("");
  const [applyingDiscount, setApplyingDiscount] = useState(false);
  const [authFields, setAuthFields] = useState({ email: "", password: "", name: "", phone: "" });
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);

  const trackVisitMut = useTrackVisit();
  useEffect(() => {
    try {
      const today = new Date().toISOString().slice(0, 10);
      const key = `ub_visit_${today}`;
      if (typeof window !== "undefined" && !window.sessionStorage.getItem(key)) {
        window.sessionStorage.setItem(key, "1");
        trackVisitMut.mutate({ data: { path: window.location.pathname || "/" } });
      }
    } catch {
      // ignore
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!customerToken) { setCustomer(null); return; }
    fetch("/api/customers/me", { headers: { Authorization: `Bearer ${customerToken}` } })
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (data) setCustomer(data);
        else {
          try { localStorage.removeItem("customer_token"); } catch {}
          setCustomerToken(null);
        }
      }).catch(() => {});
  }, [customerToken]);

  useEffect(() => {
    if (!customerToken || !customer) return;
    fetch("/api/customers/carts", { headers: { Authorization: `Bearer ${customerToken}` } })
      .then(r => r.ok ? r.json() : [])
      .then(setSavedCarts)
      .catch(() => {});
  }, [customer, customerToken]);

  const [editingProduct, setEditingProduct] = useState<number | null>(null);
  const [collapsedAisles, setCollapsedAisles] = useState<Record<string, boolean>>({});
  const [adminCategory, setAdminCategory] = useState<string>("Todas");
  const [adminProductSearch, setAdminProductSearch] = useState<string>("");

  const buyerCategoriesRef = useDragScroll();
  const buyerAislesRef = useDragScroll();
  const adminTabsRef = useDragScroll();
  const adminCategoriesRef = useDragScroll();

  const [newCategoryName, setNewCategoryName] = useState("");
  const [newAisleName, setNewAisleName] = useState("");
  const [newSubcatName, setNewSubcatName] = useState("");
  const [savedCats, setSavedCats] = useState(false);
  const [savedAisles, setSavedAisles] = useState(false);
  const [savedSubcats, setSavedSubcats] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");

  const [formState, setFormState] = useState(DEFAULT_FORM);
  const [settingsDraft, setSettingsDraft] = useState<Settings>(settings);
  const [activeSlide, setActiveSlide] = useState(0);

  useEffect(() => {
    if (menu?.settings) setSettingsDraft(menu.settings);
  }, [menu?.settings]);

  const bannerSlides: BannerSlide[] = settings.bannerSlides ?? [];
  const effectiveSlides: BannerSlide[] =
    bannerSlides.length > 0
      ? bannerSlides
      : settings.bannerImage || settings.bannerText
        ? [{ image: settings.bannerImage, title: settings.bannerText, description: settings.bannerDescription }]
        : [];

  useEffect(() => {
    if (effectiveSlides.length <= 1) return;
    const timer = setInterval(() => {
      setActiveSlide((prev) => (prev + 1) % effectiveSlides.length);
    }, 4500);
    return () => clearInterval(timer);
  }, [effectiveSlides.length]);

  useEffect(() => {
    if (categories.length > 0 && (!activeCategory || !categories.includes(activeCategory))) {
      setActiveCategory(categories[0]);
    }
  }, [categories, activeCategory]);

  useEffect(() => {
    if (!menu) return;
    if (settings.pageTitle) document.title = settings.pageTitle;
    const iconUrl = settings.favicon || settings.logo;
    if (iconUrl) {
      let link = document.querySelector("link[rel~='icon']") as HTMLLinkElement | null;
      if (!link) {
        link = document.createElement("link");
        link.rel = "icon";
        document.head.appendChild(link);
      }
      link.href = iconUrl;
    }
  }, [menu, settings.pageTitle, settings.favicon, settings.logo]);

  const filteredProducts = useMemo(() => {
    if (!debouncedSearch) {
      return products.filter((p) => p.category === activeCategory);
    }
    const q = normalize(debouncedSearch);
    // Category exact match → show ALL products in that category
    const matchedCategory = categories.find((cat) => normalize(cat).includes(q));
    if (matchedCategory) return products.filter((p) => p.category === matchedCategory);
    // Aisle exact match → show ALL products in that aisle
    const matchedAisle = aisles.find((a) => normalize(a).includes(q));
    if (matchedAisle) return products.filter((p) => p.aisle === matchedAisle);
    // Fuzzy product-level match (name only, to avoid false positives)
    return products.filter((p) => fuzzyMatch(debouncedSearch, p.name));
  }, [products, activeCategory, debouncedSearch, categories, aisles]);

  const groupedByAisle = useMemo(
    () =>
      filteredProducts.reduce<Record<string, Product[]>>((acc, product) => {
        if (!acc[product.aisle]) acc[product.aisle] = [];
        acc[product.aisle].push(product);
        return acc;
      }, {}),
    [filteredProducts],
  );
  const activeAisles = useMemo(
    () =>
      aislesData
        .map((a) => a.name)
        .filter((name) => groupedByAisle[name])
        .concat(
          Object.keys(groupedByAisle).filter(
            (name) => !aislesData.some((a) => a.name === name),
          ),
        ),
    [aislesData, groupedByAisle],
  );

  useEffect(() => {
    if (activeAisles.length === 0) {
      if (activeAisle !== "") setActiveAisle("");
      return;
    }
    if (!activeAisles.includes(activeAisle)) {
      setActiveAisle(activeAisles[0]);
    }
  }, [activeAisles, activeAisle]);

  const cartTotal = useMemo(() => cart.reduce((sum, item) => {
    const fee = (item.transferenciaEnabled && item.transferenciaAmount) ? item.transferenciaAmount : 0;
    return sum + (item.price + fee) * item.quantity;
  }, 0), [cart]);
  const cartItemCount = useMemo(() => cart.reduce((count, item) => count + item.quantity, 0), [cart]);

  // Check if the store is open based on admin-configured hours (client-side pre-check).
  // The server independently validates this on every POST /orders request.
  const isStoreOpen = (() => {
    const openTime = settings.openTime ?? "11:00";
    const closeTime = settings.closeTime ?? "23:00";
    const now = new Date();
    const currentMinutes = now.getHours() * 60 + now.getMinutes();
    const [oh = 11, om = 0] = openTime.split(":").map(Number);
    const [ch = 23, cm = 0] = closeTime.split(":").map(Number);
    const openMinutes = oh * 60 + om;
    const closeMinutes = ch * 60 + cm;
    // Cross-midnight range (e.g. 11:00 – 01:00): open if AFTER open OR BEFORE close
    if (closeMinutes <= openMinutes) {
      return currentMinutes >= openMinutes || currentMinutes < closeMinutes;
    }
    return currentMinutes >= openMinutes && currentMinutes < closeMinutes;
  })();

  const deliveryMinimum = settings.deliveryMinimum ?? 10000;
  const hasLocations = deliveryLocations.length > 0;
  const DELIVERY_COST = selectedLocation?.price ?? (hasLocations ? 0 : 3000);
  const isDelivery = checkoutForm.deliveryType === "delivery";
  const orderTotal = cartTotal + (isDelivery ? DELIVERY_COST : 0);
  const discountSavings = appliedDiscount?.discountAmount ?? 0;
  const finalTotal = cartTotal - discountSavings + (isDelivery ? DELIVERY_COST : 0);
  const deliveryMinimumMet =
    !isDelivery || cartTotal >= deliveryMinimum;

  const showToast = (message: string) => {
    setToast({ show: true, message, actionLabel: "", onAction: null });
    setTimeout(() => setToast({ show: false, message: "", actionLabel: "", onAction: null }), 3000);
  };

  const applyDiscount = async () => {
    if (!discountCode.trim()) return;
    setApplyingDiscount(true);
    setDiscountError("");
    try {
      const r = await fetch("/api/discounts/apply", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(customerToken ? { Authorization: `Bearer ${customerToken}` } : {}) },
        body: JSON.stringify({ code: discountCode.trim(), orderTotal: cartTotal }),
      });
      const data = await r.json();
      if (!r.ok) { setDiscountError(data.message ?? "Código inválido."); return; }
      setAppliedDiscount(data);
    } catch { setDiscountError("Error al validar el código."); }
    finally { setApplyingDiscount(false); }
  };

  const saveCurrentCart = async () => {
    if (!customerToken || cart.length === 0) return;
    const name = `Carrito ${new Date().toLocaleDateString("es-CL")}`;
    try {
      const r = await fetch("/api/customers/carts", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${customerToken}` },
        body: JSON.stringify({ name, items: cart }),
      });
      if (r.ok) {
        const saved = await r.json();
        setSavedCarts(prev => [...prev, saved]);
        showToast("Carrito guardado ✓");
      }
    } catch {}
  };

  const loadSavedCart = (savedCart: { items: any[] }) => {
    setCart(savedCart.items.map((item: any, i: number) => ({ ...item, cartItemId: Date.now() + i })));
    setShowAccountPanel(false);
    showToast("Carrito cargado");
  };

  const deleteSavedCart = async (cartId: number) => {
    if (!customerToken) return;
    try {
      await fetch(`/api/customers/carts/${cartId}`, { method: "DELETE", headers: { Authorization: `Bearer ${customerToken}` } });
      setSavedCarts(prev => prev.filter(c => c.id !== cartId));
    } catch {}
  };

  const handleCustomerAuth = async (e: FormEvent) => {
    e.preventDefault();
    setAuthLoading(true);
    setAuthError("");
    try {
      const url = authMode === "login" ? "/api/customers/login" : "/api/customers/register";
      const body = authMode === "login"
        ? { email: authFields.email, password: authFields.password }
        : { email: authFields.email, password: authFields.password, name: authFields.name, phone: authFields.phone };
      const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      const data = await r.json();
      if (!r.ok) { setAuthError(data.message ?? "Error al procesar."); return; }
      try { localStorage.setItem("customer_token", data.token); } catch {}
      setCustomerToken(data.token);
      setCustomer(data.customer);
      setShowAuthModal(false);
      setAuthFields({ email: "", password: "", name: "", phone: "" });
      showToast(`¡Bienvenido, ${data.customer.name}!`);
    } catch { setAuthError("Error de conexión. Intenta de nuevo."); }
    finally { setAuthLoading(false); }
  };

  const handleCustomerLogout = () => {
    try { localStorage.removeItem("customer_token"); } catch {}
    setCustomerToken(null);
    setCustomer(null);
    setSavedCarts([]);
    setShowAccountPanel(false);
    showToast("Sesión cerrada");
  };

  const compressImage = (file: File, maxPx = 1800, quality = 0.9): Promise<Blob> =>
    new Promise((resolve) => {
      if (file.size < 150_000) { resolve(file); return; }
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        URL.revokeObjectURL(url);
        let { width, height } = img;
        if (width > maxPx || height > maxPx) {
          const ratio = Math.min(maxPx / width, maxPx / height);
          width = Math.round(width * ratio);
          height = Math.round(height * ratio);
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d")!.drawImage(img, 0, 0, width, height);
        const useWebP = canvas.toDataURL("image/webp").startsWith("data:image/webp");
        canvas.toBlob(
          (blob) => resolve(blob ?? file),
          useWebP ? "image/webp" : "image/jpeg",
          quality,
        );
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(file); };
      img.src = url;
    });

  const handleImageUpload = async (
    e: ChangeEvent<HTMLInputElement>,
    callback: (url: string) => void,
    options: { skipCompression?: boolean } = {},
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      showToast(options.skipCompression ? "Subiendo..." : "Comprimiendo y subiendo...");
      const compressed: Blob = options.skipCompression ? file : await compressImage(file);
      const ext = compressed.type === "image/webp" ? "webp" : "jpg";
      const uploadName = file.name.replace(/\.[^.]+$/, "") + "." + ext;
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const presignRes = await fetch(`${apiBase}/storage/uploads/request-url`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: uploadName,
          size: compressed.size,
          contentType: compressed.type,
        }),
      });
      if (!presignRes.ok) throw new Error("presign failed");
      const { uploadURL, objectPath } = await presignRes.json();
      const putRes = await fetch(uploadURL, {
        method: "PUT",
        headers: { "Content-Type": compressed.type },
        body: compressed,
      });
      if (!putRes.ok) throw new Error("upload failed");
      const servedUrl = `${apiBase}/storage${objectPath}`;
      callback(servedUrl);
      const savings = Math.round((1 - compressed.size / file.size) * 100);
      showToast(savings > 5 ? `Imagen subida (${savings}% más liviana)` : "Imagen subida");
    } catch (err) {
      console.error(err);
      showToast("Error al subir imagen");
    } finally {
      if (e.target) e.target.value = "";
    }
  };

  const resolveImageUrl = async (
    url: string,
    setFn: (resolved: string) => void,
  ) => {
    if (!url) return;
    const isDirect =
      url.includes("/storage/objects/") ||
      /\.(jpg|jpeg|png|webp|gif|avif|svg)(\?|$)/i.test(url) ||
      url.startsWith("data:");
    if (isDirect) return;
    try {
      showToast("Resolviendo imagen...");
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(
        `${apiBase}/resolve-image?url=${encodeURIComponent(url)}`,
      );
      if (!res.ok) throw new Error("no resolved");
      const { resolved } = await res.json();
      setFn(resolved);
      showToast("Imagen lista");
    } catch {
      showToast("No se pudo resolver el enlace — prueba el link directo a la imagen");
    }
  };

  const handleDeleteStorageImage = async (
    imageUrl: string,
    clearFn: () => void,
  ) => {
    if (!imageUrl.includes("/storage/objects/")) return;
    try {
      const marker = "/storage/objects/";
      const idx = imageUrl.indexOf(marker);
      const objectPath = imageUrl.slice(idx + "/storage".length);
      const apiBase = `${import.meta.env.BASE_URL}api`;
      await fetch(`${apiBase}${objectPath}`, { method: "DELETE" });
      clearFn();
      showToast("Imagen eliminada");
    } catch {
      showToast("Error al eliminar imagen");
    }
  };

  const handleAddToCartClick = (product: Product) => {
    if (product.options && product.options.length > 0) {
      setOptionModalInfo({ product, selectedOption: product.options[0] });
    } else {
      processAddToCart(product, null);
    }
  };

  const processAddToCart = (product: Product, selectedOption: string | null) => {
    setCart((prev) => {
      const existing = prev.find(
        (item) => item.id === product.id && item.selectedOption === selectedOption,
      );
      if (existing) {
        return prev.map((item) =>
          item === existing ? { ...item, quantity: item.quantity + 1 } : item,
        );
      }
      return [
        ...prev,
        { ...product, quantity: 1, selectedOption, cartItemId: Date.now() },
      ];
    });
    setOptionModalInfo(null);
    showToast(`¡${product.name} añadido al carrito!`);
  };

  const addDeposit = (product: Product) => {
    const depositProduct: Product = {
      ...product,
      id: -product.id,
      name: `Depósito retornable - ${product.name}`,
      price: product.depositoAmount ?? 500,
      options: [],
      optionsTitle: "",
    };
    processAddToCart(depositProduct, null);
  };

  const updateQuantity = (cartItemId: number, delta: number) => {
    setCart((prev) =>
      prev
        .map((item) =>
          item.cartItemId === cartItemId
            ? { ...item, quantity: item.quantity + delta }
            : item,
        )
        .filter((item) => item.quantity > 0),
    );
  };

  const createOrderMut = useCreateOrder();

  const openCheckout = () => {
    if (cart.length === 0) return;
    if (customer) {
      setCheckoutForm(f => ({
        ...f,
        customerName: customer.name,
        phone: customer.phone ?? f.phone,
      }));
    }
    setCheckoutStep(1);
    setCheckoutOpen(true);
  };

  const submitCheckout = async (e: FormEvent) => {
    e.preventDefault();
    if (cart.length === 0) return;
    const { customerName, address, phone, notes, deliveryType } = checkoutForm;
    if (!customerName.trim() || !phone.trim()) {
      showToast("Completa nombre y teléfono");
      return;
    }
    if (deliveryType === "delivery" && !address.trim()) {
      showToast("Ingresa la dirección de entrega");
      return;
    }

    // Client-side validations before opening WhatsApp (avoid wasted messages)
    if (!isStoreOpen) {
      showToast(`Estamos fuera de nuestro horario de atención. La recepción de pedidos se habilita a las ${settings.openTime ?? "11:00"}.`);
      return;
    }
    if (!deliveryMinimumMet) {
      showToast(`El monto mínimo para Delivery es $${deliveryMinimum.toLocaleString("es-CL")}. Agrega más productos para continuar.`);
      return;
    }

    const items = cart.map((item) => ({
      name: item.name,
      quantity: item.quantity,
      price: item.price,
      selectedOption: item.selectedOption ?? undefined,
    }));

    const tipoEntrega = deliveryType === "delivery" ? "🛵 Delivery" : "🏪 Retiro en tienda";
    let text = `¡Hola! Soy *${customerName}* y quiero hacer el siguiente pedido:\n\n`;
    cart.forEach((item) => {
      const optionText = item.selectedOption ? ` (${item.selectedOption})` : "";
      text += `👉 ${item.quantity}x ${item.name}${optionText} - $${(
        item.price * item.quantity
      ).toLocaleString("es-CL")}\n`;
    });
    text += `\n📦 *Tipo de entrega:* ${tipoEntrega}`;
    if (isDelivery && selectedLocation) {
      text += ` — ${selectedLocation.name}`;
    }
    text += `\n📍 *Dirección:* ${address}`;
    if (phone) text += `\n📞 *Teléfono:* ${phone}`;
    if (notes) text += `\n📝 *Notas:* ${notes}`;
    if (isDelivery) {
      text += `\n\n🛵 *Costo delivery:* $${DELIVERY_COST.toLocaleString("es-CL")}`;
    }
    if (discountSavings > 0) {
      text += `\n🏷️ *Descuento (${appliedDiscount?.code}):* -$${discountSavings.toLocaleString("es-CL")}`;
    }
    text += `\n\n💰 *Total: $${finalTotal.toLocaleString("es-CL")}*\n\n¿Me confirmas?`;

    // Abrir WhatsApp antes del await para que quede dentro del gesto del usuario.
    // Los navegadores móviles bloquean window.open() si se llama después de un await.
    const cleanWa = (settings.whatsapp ?? "").replace(/\D/g, "");
    try {
      window.open(
        `https://wa.me/${cleanWa}?text=${encodeURIComponent(text)}`,
        "_blank",
      );
    } catch {}

    setCart([]);
    setAppliedDiscount(null);
    setDiscountCode("");
    setCheckoutOpen(false);
    setIsCartOpen(false);
    setCheckoutForm({ customerName: "", address: "", phone: "", notes: "", deliveryType: "retiro" });
    showToast("Pedido enviado");

    try {
      await createOrderMut.mutateAsync({
        data: {
          customerName,
          address,
          phone,
          notes,
          items,
          subtotal: discountSavings > 0 ? cartTotal : undefined,
          discountCode: discountSavings > 0 ? (appliedDiscount?.code ?? undefined) : undefined,
          discountAmount: discountSavings > 0 ? discountSavings : undefined,
          total: finalTotal,
          deliveryType,
        },
      });
    } catch (err) {
      console.error(err);
    }
  };

  const handleAddCategory = () => {
    const name = newCategoryName.trim();
    if (!name || categories.includes(name)) return;
    setNewCategoryName("");
    showToast("Categoría añadida");
    createCategoryMut.mutate(
      { data: { name } },
      {
        onSuccess: (newCat) =>
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, categories: [...old.categories, newCat] } : old,
          ),
        onError: refreshMenu,
      },
    );
  };
  const handleAddAisle = () => {
    const name = newAisleName.trim();
    if (!name || aisles.includes(name)) return;
    setNewAisleName("");
    showToast("Pasillo añadido");
    createAisleMut.mutate(
      { data: { name } },
      {
        onSuccess: (newAisle) =>
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, aisles: [...old.aisles, newAisle] } : old,
          ),
        onError: refreshMenu,
      },
    );
  };
  const handleAddSubcat = () => {
    const name = newSubcatName.trim();
    if (!name || subcategories.includes(name)) return;
    setNewSubcatName("");
    showToast("Subcategoría añadida");
    createSubcategoryMut.mutate(
      { data: { name } },
      {
        onSuccess: (newSub) =>
          queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
            old ? { ...old, subcategories: [...old.subcategories, newSub] } : old,
          ),
        onError: refreshMenu,
      },
    );
  };

  const handleDragStart = (e: DragEvent<HTMLDivElement>, index: number) =>
    e.dataTransfer.setData("dragIndex", String(index));
  const handleDrop = (e: DragEvent<HTMLDivElement>, dropIndex: number) => {
    const dragIndex = parseInt(e.dataTransfer.getData("dragIndex"));
    const newOrder = [...categoriesData];
    const [dragged] = newOrder.splice(dragIndex, 1);
    newOrder.splice(dropIndex, 0, dragged);
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, categories: newOrder } : old,
    );
    reorderCategoriesMut.mutate(
      { data: { ids: newOrder.map((c) => c.id) } },
      { onError: refreshMenu },
    );
  };

  const moveAisle = (index: number, dir: -1 | 1) => {
    const target = index + dir;
    if (target < 0 || target >= aislesData.length) return;
    const newOrder = [...aislesData];
    [newOrder[index], newOrder[target]] = [newOrder[target], newOrder[index]];
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, aisles: newOrder } : old,
    );
    reorderAislesMut.mutate(
      { data: { ids: newOrder.map((a) => a.id) } },
      { onError: refreshMenu },
    );
  };

  const moveSubcat = (index: number, dir: -1 | 1) => {
    const target = index + dir;
    if (target < 0 || target >= subcategoriesData.length) return;
    const newOrder = [...subcategoriesData];
    [newOrder[index], newOrder[target]] = [newOrder[target], newOrder[index]];
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, subcategories: newOrder } : old,
    );
    reorderSubcategoriesMut.mutate(
      { data: { ids: newOrder.map((s) => s.id) } },
      { onError: refreshMenu },
    );
  };

  const startEditing = (product: Product) => {
    setEditingProduct(product.id);
    setFormState({
      name: product.name,
      price: String(product.price),
      image: product.image,
      category: product.category,
      aisle: product.aisle,
      subcategory: product.subcategory || "",
      optionsTitle: product.optionsTitle || "",
      optionsString: product.options ? product.options.join(", ") : "",
      oferta: product.oferta ?? false,
      depositoEnabled: product.depositoEnabled ?? false,
      depositoAmount: String(product.depositoAmount ?? 500),
      transferenciaEnabled: product.transferenciaEnabled ?? false,
      transferenciaAmount: String(product.transferenciaAmount ?? 0),
    });
  };

  const cancelEditing = () => {
    setEditingProduct(null);
    setFormState(DEFAULT_FORM);
  };

  const handleLogin = async () => {
    try {
      const res = await adminLoginMut.mutateAsync({
        data: { password: passwordInput },
      });
      if (res.ok) {
        const role = (res.role === "delivery" ? "delivery" : "full") as "full" | "delivery";
        setAdminRole(role);
        if (role === "delivery") setAdminTab("orders");
        setView("admin");
        setPasswordInput("");
      } else {
        showToast("Contraseña incorrecta");
      }
    } catch {
      showToast("Contraseña incorrecta");
    }
  };

  const resolveImageForSave = async (url: string): Promise<string> => {
    if (!url) return url;
    const isDirect =
      url.includes("/storage/objects/") ||
      /\.(jpg|jpeg|png|webp|gif|avif|svg)(\?|$)/i.test(url) ||
      url.startsWith("data:");
    if (isDirect) return url;
    try {
      const apiBase = `${import.meta.env.BASE_URL}api`;
      const res = await fetch(`${apiBase}/resolve-image?url=${encodeURIComponent(url)}`);
      if (!res.ok) return url;
      const { resolved } = await res.json();
      return resolved || url;
    } catch {
      return url;
    }
  };

  const saveProduct = (e: FormEvent) => {
    e.preventDefault();
    if (!formState.category || !formState.aisle) {
      showToast("Por favor selecciona una categoría y un pasillo.");
      return;
    }
    const parsedOptions = formState.optionsString
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s !== "");
    const productData = {
      name: formState.name,
      price: parseInt(formState.price),
      category: formState.category,
      aisle: formState.aisle,
      subcategory: formState.subcategory,
      optionsTitle: formState.optionsTitle,
      options: parsedOptions,
      image: formState.image || "",
      oferta: formState.oferta,
      depositoEnabled: formState.depositoEnabled,
      depositoAmount: parseInt(formState.depositoAmount) || 500,
      transferenciaEnabled: formState.transferenciaEnabled,
      transferenciaAmount: parseInt(formState.transferenciaAmount) || 0,
    };
    if (editingProduct) {
      const id = editingProduct;
      queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
        old
          ? { ...old, products: old.products.map((p: any) => (p.id === id ? { ...p, ...productData } : p)) }
          : old,
      );
      showToast("Producto actualizado");
      cancelEditing();
      updateProductMut.mutate({ id, data: productData }, { onError: refreshMenu });
    } else {
      showToast("Producto creado");
      cancelEditing();
      createProductMut.mutate(
        { data: productData },
        {
          onSuccess: (newProduct) =>
            queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
              old ? { ...old, products: [...old.products, newProduct] } : old,
            ),
          onError: refreshMenu,
        },
      );
    }
  };

  const deleteProductHandler = (id: number) => {
    const product = products.find((p) => p.id === id);
    if (!product) return;
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, products: old.products.filter((p: any) => p.id !== id) } : old,
    );
    setToast({
      show: true,
      message: "Producto eliminado",
      actionLabel: "Deshacer",
      onAction: async () => {
        await createProductMut.mutateAsync({
          data: {
            name: product.name,
            price: product.price,
            category: product.category,
            aisle: product.aisle,
            subcategory: product.subcategory || "",
            optionsTitle: product.optionsTitle || "",
            options: product.options || [],
            image: product.image,
          },
        });
        refreshMenu();
      },
    });
    deleteProductMut.mutate({ id }, { onError: refreshMenu });
  };

  const deleteCategoryHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, categories: old.categories.filter((c: any) => c.id !== id) } : old,
    );
    deleteCategoryMut.mutate({ id }, { onError: refreshMenu });
  };
  const deleteAisleHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, aisles: old.aisles.filter((a: any) => a.id !== id) } : old,
    );
    deleteAisleMut.mutate({ id }, { onError: refreshMenu });
  };
  const deleteSubcategoryHandler = (id: number) => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, subcategories: old.subcategories.filter((s: any) => s.id !== id) } : old,
    );
    deleteSubcategoryMut.mutate({ id }, { onError: refreshMenu });
  };

  const saveSettings = () => {
    queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
      old ? { ...old, settings: { ...old.settings, ...settingsDraft } } : old,
    );
    showToast("Ajustes guardados");
    updateSettingsMut.mutate(
      { data: settingsDraft },
      { onError: refreshMenu },
    );
  };

  if (!ageVerified) {
    return (
      <div className="min-h-screen bg-[#141414] flex items-center justify-center p-4">
        <div className="bg-[#1a1a1a] border border-[#ffd025]/20 rounded-3xl p-8 md:p-12 max-w-sm w-full flex flex-col items-center gap-7 shadow-2xl shadow-black/70">
          <div className="w-full flex justify-center min-h-[80px] items-center">
            {isLoading || !settings.logo ? (
              <div className="w-44 h-16 rounded-2xl bg-[#252525] animate-pulse" />
            ) : (
              <img
                src={settings.logo}
                alt="Logo"
                className="max-h-24 max-w-[200px] object-contain"
                onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
              />
            )}
          </div>
          <div className="w-full h-px bg-[#ffd025]/10" />
          <div className="text-center space-y-2">
            <p className="text-gray-300 text-base leading-snug">
              Debes ser mayor de 18 años para comprar en nuestra tienda
            </p>
            <p className="text-[#ffd025] text-xl font-black mt-2">
              ¿Eres mayor de edad?
            </p>
          </div>
          <div className="flex gap-4 w-full">
            <button
              onClick={handleAgeNo}
              className="flex-1 py-3 rounded-2xl border border-white/20 text-white font-bold text-lg hover:bg-white/5 transition-colors"
            >
              No
            </button>
            <button
              onClick={handleAgeYes}
              className="flex-1 py-3 rounded-2xl bg-[#ffd025] text-[#141414] font-black text-lg hover:brightness-110 transition-all shadow-lg shadow-[#ffd025]/20"
            >
              Sí
            </button>
          </div>
          <p className="text-gray-600 text-[11px] text-center leading-relaxed">
            Al ingresar confirmas que eres mayor de 18 años y aceptas la venta responsable de alcohol.
          </p>
        </div>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="min-h-screen bg-[#141414] flex items-center justify-center">
        <div className="w-10 h-10 rounded-full border-4 border-[#ffd025]/20 border-t-[#ffd025] animate-spin" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#141414] text-[#ffd025] font-sans selection:bg-[#ffd025] selection:text-[#141414]">
      {view === "client" && (
        <div className="fixed left-0 right-0 z-40 bg-[#1a1a1a]/95 backdrop-blur-md border-t border-[#ffd025]/20 shadow-[0_-4px_24px_rgba(0,0,0,0.5)] px-4 py-3" style={{ bottom: keyboardOffset }}>
          <div className="max-w-2xl mx-auto flex items-center gap-3">
            <div className="relative flex-1">
              <Search size={17} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Buscar productos..."
                className="w-full bg-[#141414] border border-[#ffd025]/25 rounded-2xl pl-10 pr-10 py-3 text-sm text-white placeholder-gray-500 focus:outline-none focus:border-[#ffd025]/60 transition-colors"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery("")}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-white transition-colors"
                >
                  <X size={16} />
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {view === "client" && (
        <div className="fixed bottom-24 right-6 z-50 flex flex-col items-center gap-3">
          <a
            href="https://www.instagram.com/fellasmarketpm/"
            target="_blank"
            rel="noopener noreferrer"
            aria-label="Instagram"
            className="text-white p-3 rounded-full shadow-2xl hover:scale-110 active:scale-95 transition-transform border-2 border-[#141414]"
            style={{ background: "linear-gradient(135deg, #833ab4, #E1306C, #fd1d1d)" }}
          >
            <Instagram size={20} />
          </a>
          {showScrollTop && (
            <button
              onClick={scrollToTop}
              aria-label="Volver arriba"
              className="bg-[#ffd025] text-[#141414] p-3.5 rounded-full shadow-2xl shadow-[#ffd025]/40 hover:scale-110 active:scale-95 transition-transform border-2 border-[#141414]"
            >
              <ArrowUp size={22} strokeWidth={3} />
            </button>
          )}
        </div>
      )}
      {toast.show && (
        <div className="fixed top-24 right-4 z-50 bg-[#ffd025] text-[#141414] px-6 py-3 rounded-full font-bold shadow-lg shadow-[#ffd025]/20 flex items-center gap-3 animate-slide-in-right">
          <CheckCircle size={20} /> {toast.message}
          {toast.actionLabel && toast.onAction && (
            <button
              type="button"
              onClick={() => {
                toast.onAction?.();
                setToast({ show: false, message: "", actionLabel: "", onAction: null });
              }}
              className="ml-2 bg-[#141414] text-white px-3 py-1 rounded-full text-xs uppercase"
            >
              {toast.actionLabel}
            </button>
          )}
        </div>
      )}

      {view === "client" && (
        <header className="bg-[#141414] border-b border-[#ffd025]/20 shadow-lg shadow-[#ffd025]/5">
          <div className="max-w-6xl mx-auto px-4 h-16 md:h-20 flex justify-between items-center gap-3">
            <button
              className="flex-shrink-0 focus:outline-none"
              onClick={() => {
                setSearchQuery("");
                if (categories[0]) setActiveCategory(categories[0]);
                window.scrollTo({ top: 0, behavior: "smooth" });
              }}
            >
              {settings.logo ? (
                <img
                  src={settings.logo}
                  alt="Logo"
                  className="h-10 md:h-14 object-contain max-w-[110px] md:max-w-[200px]"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src =
                      "https://placehold.co/300x100/141414/ffd025?text=LOGO";
                  }}
                />
              ) : (
                <span className="text-[#ffd025] font-black text-lg">{settings.pageTitle}</span>
              )}
            </button>
            <div className="flex items-center gap-2 flex-shrink-0">
              {activeAisles.length > 1 && (
                <div className="relative">
                  <button
                    onClick={() => setShowAisleMenu((v) => !v)}
                    className="inline-flex items-center gap-1 px-2.5 py-1 text-[10px] md:text-xs font-bold uppercase rounded-full border border-[#ffd025]/30 bg-[#1a1a1a] text-[#ffd025] hover:bg-[#ffd025]/10 transition-all"
                  >
                    <Layers size={11} />
                    <span className="max-w-[70px] md:max-w-[120px] truncate">{activeAisle}</span>
                    <ChevronDown size={11} className={`transition-transform ${showAisleMenu ? "rotate-180" : ""}`} />
                  </button>
                  {showAisleMenu && (
                    <div className="absolute right-0 top-full mt-1 z-50 bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl shadow-xl py-1 min-w-[150px] animate-fade-in">
                      {activeAisles.map((aisle) => (
                        <button
                          key={aisle}
                          onClick={() => { setActiveAisle(aisle); setShowAisleMenu(false); }}
                          className={`block w-full text-left px-4 py-2 text-[11px] md:text-xs font-bold uppercase transition-all ${
                            activeAisle === aisle
                              ? "text-[#ffd025] bg-[#ffd025]/10"
                              : "text-gray-400 hover:text-[#ffd025] hover:bg-[#ffd025]/5"
                          }`}
                        >
                          {aisle}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}
              <button
                onClick={() => { if (customer) { setShowAccountPanel(p => !p); } else { setShowAuthModal(true); setAuthMode("login"); } }}
                className="relative p-2 text-[#ffd025]/60 hover:text-[#ffd025] transition-colors rounded-full hover:bg-[#ffd025]/10"
                title={customer ? customer.name : "Iniciar sesión"}
              >
                <User size={22} />
                {customer && <span className="absolute top-1.5 right-1.5 w-2 h-2 bg-green-400 rounded-full border border-[#141414]" />}
              </button>
              <button
                onClick={() => setView("admin-login")}
                className="p-2 text-[#ffd025]/60 hover:text-[#ffd025] transition-colors rounded-full hover:bg-[#ffd025]/10"
                title="Admin"
              >
                <SettingsIcon size={22} />
              </button>
              <button
                onClick={() => setIsCartOpen(true)}
                className="relative bg-[#ffd025] text-[#141414] p-2.5 md:p-3 rounded-full hover:scale-105 transition-transform shadow-lg shadow-[#ffd025]/20"
              >
                <ShoppingCart size={22} />
                {cartItemCount > 0 && (
                  <span className="absolute -top-2 -right-2 bg-red-500 text-white text-xs font-bold w-5 h-5 flex items-center justify-center rounded-full border-2 border-[#141414] animate-bounce">
                    {cartItemCount}
                  </span>
                )}
              </button>
            </div>
          </div>
        </header>
      )}

      {view === "client" && (
        <>
        {/* Business hours banner — shown only when store is outside operating hours */}
        {!isStoreOpen && (
          <div className="bg-[#1a0f00] border-b border-amber-600/40 px-4 py-3">
            <div className="max-w-6xl mx-auto flex items-start gap-3">
              <span className="text-amber-400 text-base leading-none mt-0.5 flex-shrink-0">🕐</span>
              <p className="text-[13px] text-amber-300 leading-snug">
                <span className="font-bold">Estamos fuera de nuestro horario de atención.</span>{" "}
                Puedes revisar nuestro catálogo y armar tu carrito, pero la recepción de pedidos se habilitará a las{" "}
                <span className="font-black text-amber-400">{settings.openTime ?? "11:00"}</span>.
              </p>
            </div>
          </div>
        )}
        <main className="pb-24 relative">
          {optionModalInfo && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
              <div
                className="absolute inset-0 bg-black/80 backdrop-blur-sm"
                onClick={() => setOptionModalInfo(null)}
              ></div>
              <div className="relative bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl p-6 w-full max-w-sm shadow-2xl animate-fade-in">
                <button
                  onClick={() => setOptionModalInfo(null)}
                  className="absolute top-4 right-4 text-gray-400 hover:text-white"
                >
                  <X size={24} />
                </button>
                <h3 className="text-xl font-black text-white mb-1">
                  {optionModalInfo.product.name}
                </h3>
                <p className="text-[#ffd025] font-bold mb-6">
                  ${Number(optionModalInfo.product.price || 0).toLocaleString("es-CL")}
                </p>
                <div className="mb-6">
                  <label className="block text-sm font-bold text-gray-400 uppercase mb-3">
                    {optionModalInfo.product.optionsTitle || "Elige una opción:"}
                  </label>
                  <div className="space-y-2">
                    {optionModalInfo.product.options.map((opt) => (
                      <label
                        key={opt}
                        className={`flex items-center gap-3 p-3 rounded-xl border cursor-pointer transition-colors ${
                          optionModalInfo.selectedOption === opt
                            ? "border-[#ffd025] bg-[#ffd025]/10"
                            : "border-gray-700 hover:border-gray-500"
                        }`}
                      >
                        <input
                          type="radio"
                          name="product_option"
                          value={opt}
                          checked={optionModalInfo.selectedOption === opt}
                          onChange={() =>
                            setOptionModalInfo({
                              ...optionModalInfo,
                              selectedOption: opt,
                            })
                          }
                          className="accent-[#ffd025] w-4 h-4"
                        />
                        <span className="text-white text-sm font-medium">{opt}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <button
                  onClick={() =>
                    processAddToCart(
                      optionModalInfo.product,
                      optionModalInfo.selectedOption,
                    )
                  }
                  className="w-full py-3 bg-[#ffd025] text-[#141414] rounded-xl font-bold uppercase hover:bg-[#e5b81a] transition-all shadow-lg"
                >
                  Confirmar y Añadir
                </button>
              </div>
            </div>
          )}

          {!searchQuery && effectiveSlides.length > 0 && (
            <div className="max-w-6xl mx-auto px-4 mt-6">
              <div
                className="relative w-full h-48 md:h-80 shadow-lg border border-[#ffd025]/10 bg-black select-none"
                style={{ borderRadius: "2rem", overflow: "hidden", transform: "translateZ(0)", WebkitMaskImage: "-webkit-radial-gradient(white, black)" }}
              >
                {effectiveSlides.map((slide, i) => (
                  <div
                    key={i}
                    className="absolute inset-0 transition-opacity duration-700"
                    style={{ opacity: i === activeSlide ? 1 : 0, zIndex: i === activeSlide ? 2 : 1 }}
                  >
                    {slide.image && (
                      <img
                        loading={i === 0 ? "eager" : "lazy"}
                        decoding="async"
                        fetchPriority={i === 0 ? "high" : "auto"}
                        src={slide.image}
                        alt={slide.title || "Banner"}
                        className="absolute inset-0 w-full h-full object-cover opacity-90"
                      />
                    )}
                    <div className="absolute inset-0 bg-gradient-to-r from-[#141414]/90 via-[#141414]/50 to-transparent z-10 w-full md:w-2/3"></div>
                    <div className="absolute inset-0 z-20 flex flex-col justify-center p-6 md:p-12 w-full md:w-1/2">
                      <div className="inline-flex items-center gap-2 mb-2 bg-[#141414]/50 w-fit px-3 py-1.5 rounded-full backdrop-blur-sm border border-[#ffd025]/30">
                        <span className="w-1.5 h-1.5 rounded-full bg-[#ffd025] animate-pulse"></span>
                        <span className="text-[#ffd025] font-bold text-[8px] md:text-[10px] tracking-widest uppercase">
                          Destacado
                        </span>
                      </div>
                      {slide.title && (
                        <h1
                          className="text-xs md:text-2xl lg:text-3xl uppercase leading-tight tracking-wide text-white drop-shadow-md max-w-[60%] md:max-w-none"
                          style={{ fontFamily: "'Poppins', sans-serif", fontWeight: 900 }}
                        >
                          {slide.title}
                        </h1>
                      )}
                      {slide.description && (
                        <p
                          className="text-[10px] md:text-base lg:text-lg text-white/90 mt-1 md:mt-2 leading-snug max-w-[60%] md:max-w-md drop-shadow"
                          style={{ fontFamily: "'Poppins', sans-serif", fontWeight: 200 }}
                        >
                          {slide.description}
                        </p>
                      )}
                      {settings.logo && (
                        <div className="flex items-end mt-4 md:mt-6">
                          <img
                            src={settings.logo}
                            alt="logo tienda"
                            className="h-4 md:h-7 w-auto object-contain opacity-90 drop-shadow-[0_1px_2px_rgba(0,0,0,0.9)]"
                          />
                        </div>
                      )}
                    </div>
                  </div>
                ))}
                {effectiveSlides.length > 1 && (
                  <div className="absolute bottom-3 left-1/2 -translate-x-1/2 z-30 flex items-center gap-2">
                    {effectiveSlides.map((_, i) => (
                      <button
                        key={i}
                        onClick={() => setActiveSlide(i)}
                        className={`rounded-full transition-all duration-300 ${
                          i === activeSlide
                            ? "w-5 h-2 bg-[#ffd025]"
                            : "w-2 h-2 bg-white/40 hover:bg-white/70"
                        }`}
                      />
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {(settings.announcements ?? []).length > 0 && (
            <div
              className="w-full overflow-hidden py-2 my-5 relative"
              style={{ background: "linear-gradient(to right, #7c3aed, #ec4899, #f59e0b)" }}
            >
              <div
                className="ticker-track"
                style={{ display: "flex", width: "max-content" }}
              >
                {Array.from({ length: 6 }, (_, dup) =>
                  (settings.announcements as string[]).map((text, i) => (
                    <span
                      key={`${dup}-${i}`}
                      className="text-white text-xs md:text-sm font-black uppercase tracking-wide drop-shadow-md px-3"
                    >
                      {text}
                      <span className="ml-3 text-white/40">·</span>
                    </span>
                  ))
                )}
              </div>
            </div>
          )}

          <div className="max-w-6xl mx-auto mt-3">
            {!searchQuery && categories.length > 0 && (
              <div ref={buyerCategoriesRef} className="flex overflow-x-auto gap-3 pb-4 scrollbar-hide px-4 snap-x scroll-fade-x scroll-smooth select-none">
                {categories.map((cat) => (
                  <button
                    key={cat}
                    onClick={() => setActiveCategory(cat)}
                    className={`snap-start whitespace-nowrap px-6 py-3 rounded-full font-bold uppercase tracking-wide text-sm transition-all duration-300 border-2 ${
                      activeCategory === cat
                        ? "bg-[#ffd025] text-[#141414] border-[#ffd025] scale-105"
                        : "bg-transparent text-[#ffd025] border-[#ffd025]/30 hover:border-[#ffd025]"
                    }`}
                  >
                    {cat}
                  </button>
                ))}
              </div>
            )}

            {filteredProducts.length === 0 && (
              <div className="text-center py-20 opacity-50 flex flex-col items-center">
                <Utensils size={48} className="mb-4 text-[#ffd025]/50" />
                <p className="text-xl font-bold uppercase">No se encontraron productos.</p>
              </div>
            )}


            <div className="space-y-12 px-4">
              {activeAisles
                .filter((aisle) => searchQuery || activeAisles.length <= 1 || aisle === activeAisle)
                .map((aisle) => { const prods = groupedByAisle[aisle]; return (
                <div key={aisle} id={`aisle-${aisle}`} className="animate-fade-in">
                  <h2 className="text-xl md:text-2xl font-black uppercase tracking-widest mb-4 flex items-center gap-3 border-b border-[#ffd025]/20 pb-2">
                    <span className="w-2 h-8 bg-[#ffd025] rounded-full"></span>
                    {aisle}
                  </h2>
                  {(() => {
                    const aisleObj = aislesData.find((a) => a.name === aisle);
                    if (!aisleObj?.bannerImage && !aisleObj?.bannerTitle && !aisleObj?.bannerSubtitle) return null;
                    return (
                      <div
                        className="mb-6 relative border border-[#ffd025]/20 shadow-lg"
                        style={{ borderRadius: "1rem", overflow: "hidden", transform: "translateZ(0)", WebkitMaskImage: "-webkit-radial-gradient(white, black)", background: "linear-gradient(135deg, #ff3b8a, #9d4edd, #3a86ff)" }}
                      >
                        {aisleObj?.bannerImage && (
                          <img
                            src={aisleObj.bannerImage}
                            alt={`Banner ${aisle}`}
                            loading="lazy"
                            className="w-full h-24 md:h-56 object-cover"
                          />
                        )}
                        {aisleObj?.bannerImage && (aisleObj?.bannerTitle || aisleObj?.bannerSubtitle) && (
                          <div className="absolute inset-y-0 left-0 w-3/5 md:w-2/3 bg-gradient-to-r from-black/85 via-black/55 to-transparent" />
                        )}
                        {(aisleObj?.bannerTitle || aisleObj?.bannerSubtitle || settings.logo) && (
                          <div className={`${aisleObj?.bannerImage ? "absolute inset-y-0 left-0 w-3/5 md:w-1/2" : "w-full py-5"} flex flex-col items-start justify-center text-left pl-3 pr-2 md:pl-7 md:pr-4 z-20`}>
                            {aisleObj?.bannerTitle && (
                              <span className="text-white font-black text-[11px] leading-[1.05] md:text-2xl md:leading-tight uppercase tracking-tight md:tracking-wide drop-shadow-[0_2px_4px_rgba(0,0,0,0.9)] break-words">
                                {aisleObj.bannerTitle}
                              </span>
                            )}
                            {aisleObj?.bannerSubtitle && (
                              <span className="text-white/95 font-bold text-[9px] leading-[1.1] md:text-sm md:leading-tight drop-shadow-[0_2px_4px_rgba(0,0,0,0.9)] mt-1 break-words">
                                {aisleObj.bannerSubtitle}
                              </span>
                            )}
                            {settings.logo && (
                              <div className="flex items-end mt-3 md:mt-5">
                                <img
                                  src={settings.logo}
                                  alt="logo tienda"
                                  className="h-3 md:h-5 w-auto object-contain opacity-90 drop-shadow-[0_1px_2px_rgba(0,0,0,0.9)]"
                                />
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                  {(() => {
                    const subGroups = new Map<string, typeof prods>();
                    prods.forEach((p) => {
                      const key = p.subcategory?.trim() || "";
                      if (!subGroups.has(key)) subGroups.set(key, [] as typeof prods);
                      subGroups.get(key)!.push(p);
                    });
                    return Array.from(subGroups.entries())
                      .sort(([a], [b]) => {
                        if (!a) return -1;
                        if (!b) return 1;
                        const ai = subcategoriesData.findIndex((s) => s.name === a);
                        const bi = subcategoriesData.findIndex((s) => s.name === b);
                        return (ai === -1 ? 9999 : ai) - (bi === -1 ? 9999 : bi);
                      })
                      .map(([sub, subProds]) => (
                      <div key={sub || "_none"} className="mb-10 last:mb-0">
                        {sub && (
                          <div className="flex items-center gap-3 mb-5 mt-2">
                            <div
                              className="flex items-center gap-2.5 font-black uppercase tracking-wider text-sm md:text-base px-5 py-2 rounded-full border-2 border-[#141414]"
                              style={{
                                background: "linear-gradient(to right, #ffd025, #ffb800, #ff8a00)",
                                color: "#141414",
                                boxShadow: "0 4px 16px rgba(255,208,37,0.30)",
                              }}
                            >
                              <span className="text-xl leading-none">{subcategoryEmoji(sub)}</span>
                              <span>{sub}</span>
                            </div>
                            <div className="flex-1 h-[2px] bg-gradient-to-r from-[#ffd025]/50 via-[#ffd025]/20 to-transparent rounded-full"></div>
                            <span className="text-[10px] font-bold text-[#ffd025]/60 uppercase tracking-widest">
                              {subProds.length} {subProds.length === 1 ? "ítem" : "ítems"}
                            </span>
                          </div>
                        )}
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                          {subProds.map((product) => (
                            <div
                              key={product.id}
                              className="relative bg-gradient-to-br from-[#1e1e2e] via-[#1a1a2a] to-[#141428] border border-purple-500/15 rounded-3xl overflow-hidden hover:border-[#ffd025]/40 hover:shadow-[0_0_30px_rgba(255,208,37,0.12)] transition-all duration-500 group flex flex-col"
                              style={{ transform: "translateZ(0)", WebkitMaskImage: "-webkit-radial-gradient(white, black)" }}
                            >
                              {product.options && product.options.length > 0 && (
                          <div
                            className="absolute top-3 right-3 z-10 backdrop-blur-md text-white text-[9px] font-black px-2.5 py-1 rounded-full uppercase border border-indigo-400/30 flex items-center gap-1 shadow-lg"
                            style={{ background: "linear-gradient(to right, rgba(79,70,229,0.9), rgba(147,51,234,0.9))" }}
                          >
                            ✨ Opciones
                          </div>
                        )}
                        <div className="relative h-48 overflow-hidden bg-[#1a1a2a] rounded-t-3xl">
                          <div className="absolute inset-0 animate-pulse bg-gradient-to-r from-[#1a1a2a] via-[#252535] to-[#1a1a2a]" />
                          {product.image && (
                            <img
                              loading="lazy"
                              src={product.image}
                              alt={product.name}
                              className="absolute inset-0 w-full h-full object-cover group-hover:scale-110 transition-transform duration-500 opacity-0"
                              onLoad={(e) => {
                                const el = e.currentTarget as HTMLImageElement;
                                el.style.transition = "opacity 0.15s ease";
                                el.style.opacity = "1";
                              }}
                              onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = "none"; }}
                            />
                          )}
                          <div className="absolute inset-0 bg-gradient-to-t from-[#1a1a2a] via-transparent to-transparent opacity-60 pointer-events-none" />
                          {product.bestseller && (
                            <div
                              className="absolute top-3 left-3 z-10 text-white text-[9px] font-black px-3 py-1.5 rounded-full flex items-center gap-1 uppercase tracking-wide"
                              style={{
                                background: "linear-gradient(to right, #f97316, #ec4899)",
                                boxShadow: "0 4px 14px rgba(249,115,22,0.4)",
                              }}
                            >
                              🔥 Top ventas ⭐
                            </div>
                          )}
                          {product.oferta && (
                            <div className="absolute bottom-3 left-3 z-10">
                              <div
                                className="text-white text-[10px] font-black px-3 py-1.5 rounded-full border border-white/20 flex items-center gap-1 animate-pulse"
                                style={{
                                  background: "linear-gradient(to right, #f43f5e, #ec4899, #fbbf24)",
                                  boxShadow: "0 4px 14px rgba(236,72,153,0.4)",
                                }}
                              >
                                🏷️ OFERTAZO 🔥
                              </div>
                            </div>
                          )}
                        </div>
                        <div className="p-4 flex-1 flex flex-col justify-between">
                          <div>
                            <h3 className="font-extrabold text-[15px] text-white mb-1 line-clamp-2 leading-snug tracking-tight">
                              {product.name}
                            </h3>
                            {product.options && product.options.length > 0 && (
                              <p className="text-[10px] text-purple-300/60 mb-1 line-clamp-1 font-medium">
                                📦 {product.options.join(" · ")}
                              </p>
                            )}
                          </div>
                          <div className="flex justify-between items-center mt-3 pt-3 border-t border-purple-500/10 gap-2">
                            <div className="flex flex-col min-w-0">
                              <span className="text-[10px] font-bold text-[#ffd025]/50 uppercase tracking-widest">Precio</span>
                              <span
                                className="text-xl font-black"
                                style={{
                                  background: "linear-gradient(to right, #ffd025, #ff8a00)",
                                  WebkitBackgroundClip: "text",
                                  WebkitTextFillColor: "transparent",
                                  backgroundClip: "text",
                                  color: "transparent",
                                }}
                              >
                                ${Number(product.price || 0).toLocaleString("es-CL")}
                              </span>
                              {product.transferenciaEnabled && (product.transferenciaAmount ?? 0) > 0 && (
                                <span className="text-[9px] text-blue-400/70 font-semibold mt-0.5">
                                  +${(product.transferenciaAmount ?? 0).toLocaleString("es-CL")} transferencia
                                </span>
                              )}
                            </div>
                            {(() => {
                              const hasOptions = !!(product.options && product.options.length > 0);
                              const simpleItem = !hasOptions
                                ? cart.find((i) => i.id === product.id && i.selectedOption === null)
                                : null;
                              const totalQty = cart
                                .filter((i) => i.id === product.id)
                                .reduce((s, i) => s + i.quantity, 0);
                              if (simpleItem) {
                                return (
                                  <div className="flex items-center gap-0.5 bg-gradient-to-r from-[#ffd025]/15 to-[#ff8a00]/15 rounded-2xl p-1 border border-[#ffd025]/30 flex-shrink-0">
                                    <button
                                      onClick={() => updateQuantity(simpleItem.cartItemId, -1)}
                                      className="text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] p-2 rounded-xl transition-colors"
                                      aria-label="Quitar uno"
                                    >
                                      <Minus size={16} strokeWidth={3} />
                                    </button>
                                    <span className="font-black text-[#ffd025] min-w-[24px] text-center text-base">
                                      {simpleItem.quantity}
                                    </span>
                                    <button
                                      onClick={() => updateQuantity(simpleItem.cartItemId, 1)}
                                      className="text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] p-2 rounded-xl transition-colors"
                                      aria-label="Agregar uno"
                                    >
                                      <Plus size={16} strokeWidth={3} />
                                    </button>
                                  </div>
                                );
                              }
                              return (
                                <button
                                  onClick={() => handleAddToCartClick(product)}
                                  className="relative p-3 rounded-2xl flex-shrink-0 active:scale-95 transition-transform duration-150"
                                  style={{
                                    background: "linear-gradient(135deg, #ffd025 0%, #ff8a00 100%)",
                                    color: "#141414",
                                    WebkitTapHighlightColor: "transparent",
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                  }}
                                >
                                  <Plus size={20} strokeWidth={3} color="#141414" />
                                  {totalQty > 0 && (
                                    <span
                                      className="absolute -top-1.5 -right-1.5 text-white text-[10px] font-black rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center border-2 border-[#1a1a2a] shadow-lg"
                                      style={{ background: "linear-gradient(135deg, #f43f5e, #dc2626)" }}
                                    >
                                      {totalQty}
                                    </span>
                                  )}
                                </button>
                              );
                            })()}
                          </div>
                          {product.depositoEnabled && (
                            <button
                              onClick={() => addDeposit(product)}
                              className="mt-3 w-full flex items-center justify-center gap-2 bg-gradient-to-r from-emerald-600/20 to-teal-500/20 hover:from-emerald-500 hover:to-teal-400 border border-emerald-500/30 text-emerald-300 hover:text-white text-[11px] font-black uppercase tracking-wide py-2.5 rounded-xl transition-all shadow-sm hover:shadow-emerald-500/20"
                            >
                              🍾 Depósito <span className="font-black">${(product.depositoAmount ?? 500).toLocaleString("es-CL")}</span>
                            </button>
                          )}
                        </div>
                      </div>
                          ))}
                        </div>
                      </div>
                    ));
                  })()}
                </div>
              ); })}
            </div>
          </div>

          {(settings.contactPhone || settings.contactAddress || settings.contactHours) && (
            <div className="mt-10 pt-5 border-t border-white/10 px-4">
              <div className="flex flex-wrap items-center justify-center gap-x-8 gap-y-3">
                {settings.contactPhone && (
                  <div className="flex items-center gap-2 text-gray-400">
                    <PhoneCall size={14} className="text-[#ffd025] flex-shrink-0" />
                    <span className="text-sm">{settings.contactPhone}</span>
                  </div>
                )}
                {settings.contactAddress && (
                  <div className="flex items-center gap-2 text-gray-400">
                    <MapPinIcon size={14} className="text-[#ffd025] flex-shrink-0" />
                    <span className="text-sm">{settings.contactAddress}</span>
                  </div>
                )}
                {settings.contactHours && (
                  <div className="flex items-center gap-2 text-gray-400">
                    <ClockIcon size={14} className="text-[#ffd025] flex-shrink-0" />
                    <span className="text-sm">{settings.contactHours}</span>
                  </div>
                )}
              </div>
            </div>
          )}

          <footer className="mt-6 py-4 border-t border-gray-800/30 flex items-center justify-center gap-2 opacity-40 hover:opacity-100 transition-opacity">
            <p className="text-[10px] text-gray-600 uppercase tracking-widest font-medium m-0">
              Desarrollado por {settings.agencyName || "Nosotros"}
            </p>
            {settings.agencyLogo && (
              <img
                src={settings.agencyLogo}
                alt="Logo Agencia"
                className="h-3 md:h-4 object-contain grayscale hover:grayscale-0 transition-all duration-300 opacity-70 hover:opacity-100"
              />
            )}
          </footer>
        </main>
        </>
      )}

      {isCartOpen && (
        <div className="fixed inset-0 z-50 flex justify-end">
          <div
            className="absolute inset-0 bg-black/70 backdrop-blur-sm transition-opacity"
            onClick={() => setIsCartOpen(false)}
          ></div>
          <div className="relative w-full max-w-md bg-[#141414] border-l border-[#ffd025]/20 h-full flex flex-col shadow-2xl animate-slide-in-right">
            <div className="p-6 border-b border-[#ffd025]/20 flex justify-between items-center bg-[#1a1a1a]">
              <h2 className="text-xl font-black uppercase tracking-wider flex items-center gap-3">
                <ShoppingCart className="text-[#ffd025]" /> Tu Pedido
              </h2>
              <button
                onClick={() => setIsCartOpen(false)}
                className="p-2 text-[#ffd025]/60 hover:text-[#ffd025] rounded-full hover:bg-[#ffd025]/10"
              >
                <X size={24} />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-3">
              {cart.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-full text-center opacity-50 space-y-4">
                  <ShoppingCart size={48} className="text-[#ffd025]/50" />
                  <p className="text-lg font-bold">Carrito Vacío</p>
                </div>
              ) : (
                cart.map((item) => (
                  <div
                    key={item.cartItemId}
                    className="flex gap-4 items-center bg-[#1a1a1a] p-3 rounded-2xl border border-[#ffd025]/10"
                  >
                    {item.image ? (
                      <img
                        src={item.image}
                        alt={item.name}
                        className="w-16 h-16 rounded-xl object-cover"
                        onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = "none"; }}
                      />
                    ) : (
                      <div className="w-16 h-16 rounded-xl bg-[#1a1a1a] border border-gray-700 flex-shrink-0" />
                    )}
                    <div className="flex-1">
                      <h4 className="font-bold text-white text-sm line-clamp-1">
                        {item.name}
                      </h4>
                      {item.selectedOption && (
                        <p className="text-[10px] text-gray-400 bg-gray-800 w-fit px-1.5 py-0.5 rounded mt-1 line-clamp-1">
                          {item.selectedOption}
                        </p>
                      )}
                      <p className="text-[#ffd025] font-black mt-1">
                        ${(Number(item.price || 0) * item.quantity).toLocaleString("es-CL")}
                      </p>
                    </div>
                    <div className="flex items-center gap-2 bg-[#141414] rounded-full p-1 border border-[#ffd025]/20">
                      <button
                        onClick={() => updateQuantity(item.cartItemId, -1)}
                        className="w-8 h-8 flex items-center justify-center rounded-full hover:bg-[#ffd025]/20 text-[#ffd025] font-bold"
                      >
                        -
                      </button>
                      <span className="font-bold w-4 text-center text-white text-sm">
                        {item.quantity}
                      </span>
                      <button
                        onClick={() => updateQuantity(item.cartItemId, 1)}
                        className="w-8 h-8 flex items-center justify-center rounded-full hover:bg-[#ffd025]/20 text-[#ffd025] font-bold"
                      >
                        +
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
            {cart.length > 0 && (
              <div className="p-6 border-t border-[#ffd025]/20 bg-[#1a1a1a]">
                <div className="flex justify-between items-center mb-4">
                  <span className="text-gray-400 font-bold uppercase text-sm">Total</span>
                  <span className="text-3xl font-black text-[#ffd025]">
                    ${cartTotal.toLocaleString("es-CL")}
                  </span>
                </div>
                {/* Store closed warning */}
                {!isStoreOpen && (
                  <div className="mb-3 px-3 py-2 rounded-xl bg-amber-950/40 border border-amber-600/30 text-xs text-amber-300 text-center">
                    🕐 Pedidos habilitados desde las {settings.openTime ?? "11:00"}
                  </div>
                )}
                <button
                  onClick={openCheckout}
                  disabled={!isStoreOpen}
                  className="w-full py-4 bg-[#ffd025] text-[#141414] rounded-2xl font-black uppercase tracking-wider hover:bg-[#e5b81a] transition-all flex items-center justify-center gap-2 text-lg shadow-[0_0_20px_rgba(255,208,37,0.3)] hover:shadow-[0_0_30px_rgba(255,208,37,0.5)] disabled:opacity-40 disabled:cursor-not-allowed disabled:shadow-none"
                >
                  <Phone size={22} /> Enviar Pedido
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {checkoutOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4">
          <div className="absolute inset-0 bg-black/80 backdrop-blur-sm" onClick={() => setCheckoutOpen(false)} />
          <div className="relative w-full max-w-sm bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl shadow-2xl animate-fade-in overflow-hidden">
            {/* Header */}
            <div className="flex items-center justify-between px-5 pt-5 pb-3">
              <div>
                <p className="text-[10px] font-bold text-gray-500 uppercase tracking-widest mb-0.5">
                  Paso {checkoutStep} de 3
                </p>
                <h3 className="text-base font-black uppercase tracking-wide">
                  {checkoutStep === 1 && "¿Cómo lo recibís?"}
                  {checkoutStep === 2 && "Tus datos"}
                  {checkoutStep === 3 && "Resumen y envío"}
                </h3>
              </div>
              <button type="button" onClick={() => setCheckoutOpen(false)} className="text-gray-500 hover:text-white transition-colors ml-4 shrink-0">
                <X size={20} />
              </button>
            </div>
            {/* Progress bar */}
            <div className="px-5 pb-4">
              <div className="flex gap-1.5">
                {[1, 2, 3].map(s => (
                  <div key={s} className={`h-1 flex-1 rounded-full transition-all duration-300 ${s <= checkoutStep ? "bg-[#ffd025]" : "bg-gray-800"}`} />
                ))}
              </div>
            </div>

            <form onSubmit={submitCheckout}>
              <div className="px-5 pb-5">

                {/* ── PASO 1: Tipo de entrega ── */}
                {checkoutStep === 1 && (
                  <div className="space-y-4 animate-fade-in">
                    <div className="grid grid-cols-2 gap-3">
                      {(["retiro", "delivery"] as const).map(type => (
                        <button
                          key={type}
                          type="button"
                          onClick={() => setCheckoutForm({ ...checkoutForm, deliveryType: type })}
                          className={`flex flex-col items-center justify-center gap-2 py-5 px-3 rounded-2xl border-2 font-bold text-sm transition-all ${
                            checkoutForm.deliveryType === type
                              ? "bg-[#ffd025]/10 border-[#ffd025] text-[#ffd025]"
                              : "bg-[#141414] border-gray-700 text-gray-400 hover:border-[#ffd025]/40"
                          }`}
                        >
                          <span className="text-3xl">{type === "retiro" ? "🏪" : "🛵"}</span>
                          <span>{type === "retiro" ? "Retiro en local" : "Delivery"}</span>
                        </button>
                      ))}
                    </div>
                    {checkoutForm.deliveryType === "delivery" && !deliveryMinimumMet && (
                      <p className="text-[11px] text-amber-400 bg-amber-950/40 border border-amber-600/30 rounded-xl px-3 py-2.5 leading-snug">
                        ⚠️ Mínimo para delivery: <span className="font-black">${deliveryMinimum.toLocaleString("es-CL")}</span>. Agrega más productos.
                      </p>
                    )}
                    {checkoutForm.deliveryType === "delivery" && hasLocations && (
                      <div className="space-y-2">
                        <p className="text-[11px] font-bold text-gray-400 uppercase">Zona de envío</p>
                        <div className="space-y-2">
                          {deliveryLocations.filter(l => l.active).map(loc => (
                            <button
                              key={loc.id}
                              type="button"
                              onClick={() => setSelectedLocation(loc)}
                              className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border-2 font-semibold text-sm transition-all ${
                                selectedLocation?.id === loc.id
                                  ? "bg-[#ffd025]/10 border-[#ffd025] text-[#ffd025]"
                                  : "bg-[#141414] border-gray-700 text-gray-300 hover:border-[#ffd025]/40"
                              }`}
                            >
                              <span>📍 {loc.name}</span>
                              <span className="font-black">${loc.price.toLocaleString("es-CL")}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                    <div className="flex items-center gap-2 rounded-xl px-3 py-2.5 border border-[#ffd025]/20" style={{ background: "linear-gradient(135deg,rgba(255,208,37,0.08),rgba(255,138,0,0.06))" }}>
                      <span className="text-base">💳</span>
                      <p className="text-[11px] text-[#ffd025]/80 font-semibold leading-snug">
                        Pago por <span className="font-black text-[#ffd025]">transferencia</span>. Los datos se coordinan por WhatsApp.
                      </p>
                    </div>
                    <button
                      type="button"
                      disabled={
                        (checkoutForm.deliveryType === "delivery" && !deliveryMinimumMet) ||
                        (checkoutForm.deliveryType === "delivery" && hasLocations && !selectedLocation)
                      }
                      onClick={() => setCheckoutStep(customer ? 3 : 2)}
                      className="w-full py-3.5 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      Siguiente
                    </button>
                  </div>
                )}

                {/* ── PASO 2: Nombre y teléfono ── */}
                {checkoutStep === 2 && (
                  <div className="space-y-4 animate-fade-in">
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Nombre *</label>
                      <input
                        type="text"
                        required
                        autoFocus
                        value={checkoutForm.customerName}
                        onChange={e => setCheckoutForm({ ...checkoutForm, customerName: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                        placeholder="Tu nombre"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Teléfono *</label>
                      <input
                        type="tel"
                        required
                        value={checkoutForm.phone}
                        onChange={e => setCheckoutForm({ ...checkoutForm, phone: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                        placeholder="+56 9..."
                      />
                    </div>
                    <div className="flex gap-2 pt-1">
                      <button type="button" onClick={() => setCheckoutStep(1)} className="px-5 py-3 bg-[#141414] border border-gray-700 text-gray-400 rounded-2xl font-bold text-sm hover:border-gray-500 transition-all">
                        Atrás
                      </button>
                      <button
                        type="button"
                        disabled={!checkoutForm.customerName.trim() || !checkoutForm.phone.trim()}
                        onClick={() => setCheckoutStep(3)}
                        className="flex-1 py-3 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        Siguiente
                      </button>
                    </div>
                  </div>
                )}

                {/* ── PASO 3: Dirección, comentario, descuento y total ── */}
                {checkoutStep === 3 && (
                  <div className="space-y-3 animate-fade-in">
                    {/* Logged-in customer info chip */}
                    {customer && (
                      <div className="flex items-center gap-2.5 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-3 py-2.5">
                        <div className="w-7 h-7 rounded-full bg-[#ffd025]/15 flex items-center justify-center shrink-0">
                          <User size={14} className="text-[#ffd025]" />
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-bold text-white leading-tight truncate">{customer.name}</p>
                          <p className="text-xs text-gray-500 truncate">{customer.phone || customer.email}</p>
                        </div>
                      </div>
                    )}
                    {/* Phone field for logged-in customers without a saved phone */}
                    {customer && !checkoutForm.phone.trim() && (
                      <div>
                        <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Teléfono *</label>
                        <input
                          type="tel"
                          required
                          value={checkoutForm.phone}
                          onChange={e => setCheckoutForm({ ...checkoutForm, phone: e.target.value })}
                          className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                          placeholder="+56 9..."
                        />
                      </div>
                    )}
                    {isDelivery && (
                      <div>
                        <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Dirección de entrega *</label>
                        <input
                          type="text"
                          required={isDelivery}
                          value={checkoutForm.address}
                          onChange={e => setCheckoutForm({ ...checkoutForm, address: e.target.value })}
                          className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                          placeholder="Calle, número, comuna"
                        />
                      </div>
                    )}
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5">Comentarios (opcional)</label>
                      <textarea
                        rows={2}
                        value={checkoutForm.notes}
                        onChange={e => setCheckoutForm({ ...checkoutForm, notes: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025] resize-none"
                        placeholder="Instrucciones especiales..."
                      />
                    </div>
                    {/* Discount code */}
                    <div>
                      <label className="block text-[11px] font-bold text-gray-400 uppercase mb-1.5 flex items-center gap-1.5">
                        <Tag size={11} /> Código de descuento
                      </label>
                      {!checkoutForm.phone.trim() ? (
                        <p className="text-[11px] text-gray-500 bg-[#141414] border border-gray-800 rounded-xl px-3 py-2.5">
                          Ingresa tu teléfono para aplicar descuentos
                        </p>
                      ) : appliedDiscount ? (
                        <div className="flex items-center justify-between bg-green-900/30 border border-green-500/30 rounded-xl px-3 py-2">
                          <div className="flex items-center gap-2">
                            <CheckCircle size={13} className="text-green-400 shrink-0" />
                            <span className="text-green-400 text-sm font-bold">{appliedDiscount.label}</span>
                          </div>
                          <button type="button" onClick={() => { setAppliedDiscount(null); setDiscountCode(""); }} className="text-gray-500 hover:text-gray-300 text-xs ml-2 shrink-0">Quitar</button>
                        </div>
                      ) : (
                        <div className="flex gap-2 items-center">
                          <input
                            type="text"
                            value={discountCode}
                            onChange={e => { setDiscountCode(e.target.value.toUpperCase()); setDiscountError(""); }}
                            className="min-w-0 flex-1 bg-[#141414] border border-[#ffd025]/20 rounded-xl py-2 px-3 text-white text-sm uppercase tracking-widest focus:outline-none focus:border-[#ffd025] placeholder-gray-600"
                            placeholder="CÓDIGO"
                          />
                          <button type="button" onClick={applyDiscount} disabled={applyingDiscount || !discountCode.trim()}
                            className="px-3 py-2 bg-[#ffd025]/10 border border-[#ffd025]/20 text-[#ffd025] rounded-xl text-xs font-bold hover:bg-[#ffd025]/20 transition-colors disabled:opacity-40 shrink-0 whitespace-nowrap">
                            {applyingDiscount ? "…" : "Aplicar"}
                          </button>
                        </div>
                      )}
                      {discountError && <p className="text-red-400 text-xs mt-1">{discountError}</p>}
                    </div>
                    {/* Total */}
                    <div className="rounded-2xl bg-[#141414] border border-[#ffd025]/15 px-4 py-3 space-y-1.5">
                      <div className="flex justify-between text-sm text-gray-400">
                        <span>Subtotal</span>
                        <span>${cartTotal.toLocaleString("es-CL")}</span>
                      </div>
                      {isDelivery && (
                        <>
                          <div className="flex justify-between text-sm text-gray-400">
                            <span>Delivery</span>
                            <span>${DELIVERY_COST.toLocaleString("es-CL")}</span>
                          </div>
                          <p className="text-[10px] text-amber-400/70 leading-snug">🏘️ Tarifa fija dentro de Alerce.</p>
                        </>
                      )}
                      {discountSavings > 0 && (
                        <div className="flex justify-between text-sm text-green-400">
                          <span>Descuento ({appliedDiscount?.code})</span>
                          <span>−${discountSavings.toLocaleString("es-CL")}</span>
                        </div>
                      )}
                      <div className="flex justify-between font-black text-base text-[#ffd025] pt-1.5 border-t border-[#ffd025]/15">
                        <span>Total</span>
                        <span>${finalTotal.toLocaleString("es-CL")}</span>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <button type="button" onClick={() => setCheckoutStep(customer ? 1 : 2)} className="px-5 py-3 bg-[#141414] border border-gray-700 text-gray-400 rounded-2xl font-bold text-sm hover:border-gray-500 transition-all shrink-0">
                        Atrás
                      </button>
                      <button
                        type="submit"
                        disabled={createOrderMut.isPending || !isStoreOpen || !deliveryMinimumMet || (isDelivery && !checkoutForm.address.trim())}
                        className="flex-1 py-3 bg-[#ffd025] text-[#141414] rounded-2xl font-black text-sm hover:bg-[#e5b81a] transition-all flex items-center justify-center gap-1.5 shadow-[0_0_20px_rgba(255,208,37,0.25)] disabled:opacity-40 disabled:cursor-not-allowed disabled:shadow-none"
                      >
                        <Phone size={14} /> Confirmar pedido
                      </button>
                    </div>
                  </div>
                )}

              </div>
            </form>
          </div>
        </div>
      )}

      {/* Customer auth modal */}
      {showAuthModal && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4">
          <div className="absolute inset-0 bg-black/80 backdrop-blur-sm" onClick={() => { setShowAuthModal(false); setAuthError(""); }} />
          <div className="relative w-full max-w-sm bg-[#1a1a1a] border border-[#ffd025]/30 rounded-3xl p-6 shadow-2xl animate-fade-in">
            <div className="flex justify-between items-center mb-5">
              <h2 className="text-xl font-black uppercase">
                {authMode === "login" ? "Iniciar sesión" : "Crear cuenta"}
              </h2>
              <button onClick={() => { setShowAuthModal(false); setAuthError(""); }} className="text-gray-500 hover:text-white transition-colors">
                <X size={20} />
              </button>
            </div>
            <form onSubmit={handleCustomerAuth} className="space-y-3">
              {authMode === "register" && (
                <div>
                  <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Nombre</label>
                  <input type="text" required value={authFields.name} onChange={e => setAuthFields(f => ({ ...f, name: e.target.value }))}
                    className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="Tu nombre" />
                </div>
              )}
              <div>
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Correo electrónico</label>
                <input type="email" required value={authFields.email} onChange={e => setAuthFields(f => ({ ...f, email: e.target.value }))}
                  className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="correo@ejemplo.com" />
              </div>
              {authMode === "register" && (
                <div>
                  <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Teléfono (opcional)</label>
                  <input type="tel" value={authFields.phone} onChange={e => setAuthFields(f => ({ ...f, phone: e.target.value }))}
                    className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="+56 9..." />
                </div>
              )}
              <div>
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Contraseña</label>
                <input type="password" required minLength={authMode === "register" ? 6 : 1} value={authFields.password} onChange={e => setAuthFields(f => ({ ...f, password: e.target.value }))}
                  className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-xl p-3 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder={authMode === "register" ? "Mínimo 6 caracteres" : "••••••"} />
              </div>
              {authError && <p className="text-red-400 text-sm text-center py-1">{authError}</p>}
              <button type="submit" disabled={authLoading}
                className="w-full py-3.5 bg-[#ffd025] text-[#141414] rounded-2xl font-black uppercase tracking-wider hover:bg-[#e5b81a] transition-all disabled:opacity-50 flex items-center justify-center gap-2 mt-1">
                {authLoading ? <Loader2 size={20} className="animate-spin" /> : (authMode === "login" ? "Ingresar" : "Crear cuenta")}
              </button>
            </form>
            <p className="text-center text-sm text-gray-500 mt-4">
              {authMode === "login" ? "¿No tienes cuenta?" : "¿Ya tienes cuenta?"}
              <button onClick={() => { setAuthMode(authMode === "login" ? "register" : "login"); setAuthError(""); }} className="text-[#ffd025] font-bold ml-1 hover:underline">
                {authMode === "login" ? "Crear cuenta" : "Iniciar sesión"}
              </button>
            </p>
          </div>
        </div>
      )}

      {/* Customer account panel */}
      {showAccountPanel && customer && (
        <div className="fixed inset-0 z-[65]" onClick={() => setShowAccountPanel(false)}>
          <div className="absolute top-[72px] right-4 w-80 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-2xl shadow-2xl overflow-hidden animate-fade-in" onClick={e => e.stopPropagation()}>
            <div className="p-4 border-b border-gray-800">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-[#ffd025]/20 rounded-full flex items-center justify-center shrink-0">
                  <User size={20} className="text-[#ffd025]" />
                </div>
                <div className="min-w-0">
                  <p className="font-bold text-sm truncate">{customer.name}</p>
                  <p className="text-gray-400 text-xs truncate">{customer.email}</p>
                </div>
              </div>
            </div>
            <div className="p-4">
              <div className="flex items-center justify-between mb-2.5">
                <h3 className="text-xs font-bold text-gray-400 uppercase tracking-wider">Carritos guardados</h3>
                {cart.length > 0 && (
                  <button onClick={saveCurrentCart} className="text-xs text-[#ffd025]/80 hover:text-[#ffd025] font-bold transition-colors flex items-center gap-1">
                    <Bookmark size={11} /> Guardar actual
                  </button>
                )}
              </div>
              {savedCarts.length === 0 ? (
                <p className="text-gray-600 text-xs text-center py-4">No tienes carritos guardados.</p>
              ) : (
                <div className="space-y-2 max-h-52 overflow-y-auto">
                  {savedCarts.map(c => (
                    <div key={c.id} className="flex items-center justify-between bg-[#141414] rounded-xl px-3 py-2.5">
                      <div className="min-w-0 mr-2">
                        <p className="text-sm font-bold truncate">{c.name}</p>
                        <p className="text-xs text-gray-500">{c.items.length} producto{c.items.length !== 1 ? "s" : ""}</p>
                      </div>
                      <div className="flex gap-1 shrink-0">
                        <button onClick={() => loadSavedCart(c)} className="text-xs px-2.5 py-1.5 bg-[#ffd025]/10 text-[#ffd025] rounded-lg hover:bg-[#ffd025]/20 transition-colors font-bold">
                          Cargar
                        </button>
                        <button onClick={() => deleteSavedCart(c.id)} className="p-1.5 text-red-400/50 hover:text-red-400 rounded-lg hover:bg-red-900/20 transition-colors">
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div className="p-3 border-t border-gray-800">
              <button onClick={handleCustomerLogout} className="w-full flex items-center justify-center gap-2 py-2.5 text-sm text-red-400 hover:text-red-300 hover:bg-red-900/20 rounded-xl transition-colors font-bold">
                <LogOut size={15} /> Cerrar sesión
              </button>
            </div>
          </div>
        </div>
      )}

      {view === "admin-login" && (
        <div className="min-h-screen flex items-center justify-center p-4 bg-[#141414]">
          <div className="bg-[#1a1a1a] p-8 rounded-3xl w-full max-w-sm border border-[#ffd025]/20 shadow-2xl text-center">
            <div className="w-20 h-20 bg-[#ffd025]/10 rounded-full flex items-center justify-center mx-auto mb-6">
              <SettingsIcon className="w-10 h-10 text-[#ffd025]" />
            </div>
            <h2 className="text-2xl font-black uppercase mb-2">Administración</h2>
            <p className="text-gray-400 text-sm mb-6">Ingresa tu contraseña para acceder.</p>
            <div className="mb-6">
              <input
                type="password"
                placeholder="Contraseña"
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
                className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-4 text-center text-white text-lg tracking-widest focus:outline-none focus:border-[#ffd025] transition-colors"
              />
            </div>
            <button
              onClick={handleLogin}
              className="w-full py-4 bg-[#ffd025] text-[#141414] rounded-xl font-bold uppercase hover:bg-[#e5b81a] transition-colors mb-4 shadow-lg shadow-[#ffd025]/20"
            >
              Ingresar al Panel
            </button>
            <button
              onClick={() => {
                setView("client");
                setPasswordInput("");
              }}
              className="w-full py-3 bg-transparent text-gray-500 font-bold uppercase hover:text-white transition-colors text-sm"
            >
              Volver a la tienda
            </button>
          </div>
        </div>
      )}

      {view === "admin" && (
        <div className="min-h-screen bg-[#141414] pb-20">
          <header className="bg-[#1a1a1a] border-b border-[#ffd025]/20 sticky top-0 z-30 shadow-lg">
            <div className="max-w-6xl mx-auto px-4 h-16 flex justify-between items-center">
              <h1 className="font-black uppercase tracking-wider flex items-center gap-2">
                <SettingsIcon className="text-[#ffd025]" /> Panel Admin
              </h1>
              <button
                onClick={() => setView("client")}
                className="flex items-center gap-2 px-4 py-2 bg-[#ffd025]/10 text-[#ffd025] rounded-full font-bold hover:bg-[#ffd025] hover:text-[#141414] transition-colors text-sm uppercase"
              >
                <LogOut size={16} /> Ver Tienda
              </button>
            </div>
            <div ref={adminTabsRef} className="max-w-6xl mx-auto px-4 flex gap-4 overflow-x-auto scrollbar-hide border-t border-gray-800 scroll-fade-x scroll-smooth select-none">
              {(["products", "classifications", "orders", "stats", "settings", "social", "customers"] as const)
                .filter((tab) => adminRole === "delivery" ? tab === "orders" : true)
                .map((tab) => (
                <button
                  key={tab}
                  onClick={() => setAdminTab(tab)}
                  className={`py-3 px-4 font-bold uppercase text-sm border-b-2 transition-colors whitespace-nowrap ${
                    adminTab === tab
                      ? "border-[#ffd025] text-[#ffd025]"
                      : "border-transparent text-gray-500 hover:text-gray-300"
                  }`}
                >
                  {tab === "products"
                    ? "Productos"
                    : tab === "classifications"
                      ? "Clasificaciones"
                      : tab === "orders"
                        ? "Pedidos"
                        : tab === "stats"
                          ? "Estadísticas"
                          : tab === "social"
                            ? "Community"
                            : tab === "customers"
                              ? "Clientes"
                              : "Ajustes Generales"}
                </button>
              ))}
            </div>
          </header>

          <div className="max-w-6xl mx-auto p-4 mt-6">
            {adminTab === "orders" && <OrdersAdminPanel role={adminRole ?? "full"} />}

            {adminTab === "stats" && <StatsAdminPanel />}
            {adminTab === "customers" && <CustomersAdminPanel />}

            {adminTab === "settings" && (
              <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20 max-w-2xl animate-fade-in">
                <h2 className="text-xl font-black uppercase mb-6 flex items-center gap-2">
                  <Edit3 size={20} className="text-[#ffd025]" /> Personalización
                </h2>
                <div className="space-y-6">
                  <div className="space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      Pestaña del Navegador
                    </h3>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Título de la página
                      </label>
                      <input
                        type="text"
                        value={settingsDraft.pageTitle || ""}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, pageTitle: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                        placeholder="Ej: Urban Bite - Delivery"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Ícono de la pestaña (Favicon)
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={settingsDraft.favicon || ""}
                          onChange={(e) =>
                            setSettingsDraft({ ...settingsDraft, favicon: e.target.value })
                          }
                          onBlur={(e) =>
                            resolveImageUrl(e.target.value, (r) =>
                              setSettingsDraft((p) => ({ ...p, favicon: r })),
                            )
                          }
                          className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                          placeholder="URL o sube una imagen cuadrada"
                        />
                        <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                          <Upload size={20} />
                          <input
                            type="file"
                            accept="image/*"
                            className="hidden"
                            onChange={(e) =>
                              handleImageUpload(e, (url) =>
                                setSettingsDraft({ ...settingsDraft, favicon: url }),
                              )
                            }
                          />
                        </label>
                        {(settingsDraft.favicon || "").includes("/storage/objects/") && (
                          <button
                            type="button"
                            title="Eliminar imagen del servidor"
                            onClick={() =>
                              handleDeleteStorageImage(settingsDraft.favicon || "", () =>
                                setSettingsDraft((p) => ({ ...p, favicon: "" })),
                              )
                            }
                            className="bg-red-500/10 text-red-400 px-4 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30"
                          >
                            <Trash2 size={20} />
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="pt-4 border-t border-gray-800 space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      Diseño y Banner
                    </h3>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Logo de la tienda
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={settingsDraft.logo}
                          onChange={(e) =>
                            setSettingsDraft({ ...settingsDraft, logo: e.target.value })
                          }
                          onBlur={(e) =>
                            resolveImageUrl(e.target.value, (r) =>
                              setSettingsDraft((p) => ({ ...p, logo: r })),
                            )
                          }
                          className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                          placeholder="URL de la imagen"
                        />
                        <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                          <Upload size={20} />
                          <input
                            type="file"
                            accept="image/*"
                            className="hidden"
                            onChange={(e) =>
                              handleImageUpload(e, (url) =>
                                setSettingsDraft({ ...settingsDraft, logo: url }),
                              )
                            }
                          />
                        </label>
                        {settingsDraft.logo.includes("/storage/objects/") && (
                          <button
                            type="button"
                            title="Eliminar imagen del servidor"
                            onClick={() =>
                              handleDeleteStorageImage(settingsDraft.logo, () =>
                                setSettingsDraft((p) => ({ ...p, logo: "" })),
                              )
                            }
                            className="bg-red-500/10 text-red-400 px-4 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30"
                          >
                            <Trash2 size={20} />
                          </button>
                        )}
                      </div>
                    </div>

                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <label className="block text-xs font-bold text-gray-400 uppercase">
                          Slides del Banner
                        </label>
                        <button
                          type="button"
                          onClick={() => {
                            const newSlide: BannerSlide = { image: "", title: "", description: "" };
                            setSettingsDraft({
                              ...settingsDraft,
                              bannerSlides: [...(settingsDraft.bannerSlides ?? []), newSlide],
                            });
                          }}
                          className="flex items-center gap-1.5 text-xs bg-[#ffd025]/10 text-[#ffd025] border border-[#ffd025]/30 rounded-lg px-3 py-1.5 hover:bg-[#ffd025]/20 transition-colors font-bold"
                        >
                          <Plus size={13} /> Añadir slide
                        </button>
                      </div>
                      {(settingsDraft.bannerSlides ?? []).length === 0 && (
                        <p className="text-xs text-gray-500 italic">
                          Sin slides — se usará la imagen de banner legacy si existe.
                        </p>
                      )}
                      {(settingsDraft.bannerSlides ?? []).map((slide, idx) => (
                        <div
                          key={idx}
                          className="rounded-xl border border-[#ffd025]/20 bg-[#0d0d0d] p-3 space-y-2"
                        >
                          <div className="flex items-center justify-between mb-1">
                            <span className="text-[10px] font-bold text-[#ffd025]/60 uppercase tracking-widest">
                              Slide {idx + 1}
                            </span>
                            <button
                              type="button"
                              onClick={() => {
                                const slides = [...(settingsDraft.bannerSlides ?? [])];
                                slides.splice(idx, 1);
                                setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                              }}
                              className="text-red-400 hover:text-red-300 text-xs"
                            >
                              Eliminar
                            </button>
                          </div>
                          <div className="flex gap-2">
                            <input
                              type="text"
                              value={slide.image}
                              onChange={(e) => {
                                const slides = [...(settingsDraft.bannerSlides ?? [])];
                                slides[idx] = { ...slides[idx]!, image: e.target.value };
                                setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                              }}
                              onBlur={(e) =>
                                resolveImageUrl(e.target.value, (r) => {
                                  const slides = [...(settingsDraft.bannerSlides ?? [])];
                                  slides[idx] = { ...slides[idx]!, image: r };
                                  setSettingsDraft((p) => ({ ...p, bannerSlides: slides }));
                                })
                              }
                              className="flex-1 bg-[#141414] border border-[#ffd025]/20 rounded-lg p-2 text-white text-xs focus:border-[#ffd025]"
                              placeholder="URL imagen"
                            />
                            <label className="bg-[#ffd025]/10 text-[#ffd025] px-3 rounded-lg flex items-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/20">
                              <Upload size={15} />
                              <input
                                type="file"
                                accept="image/*"
                                className="hidden"
                                onChange={(e) =>
                                  handleImageUpload(e, (url) => {
                                    const slides = [...(settingsDraft.bannerSlides ?? [])];
                                    slides[idx] = { ...slides[idx]!, image: url };
                                    setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                                  })
                                }
                              />
                            </label>
                            {slide.image.includes("/storage/objects/") && (
                              <button
                                type="button"
                                title="Eliminar imagen del servidor"
                                onClick={() =>
                                  handleDeleteStorageImage(slide.image, () => {
                                    const slides = [...(settingsDraft.bannerSlides ?? [])];
                                    slides[idx] = { ...slides[idx]!, image: "" };
                                    setSettingsDraft((p) => ({ ...p, bannerSlides: slides }));
                                  })
                                }
                                className="bg-red-500/10 text-red-400 px-2 rounded-lg flex items-center hover:bg-red-500/20 border border-red-500/30"
                              >
                                <Trash2 size={15} />
                              </button>
                            )}
                          </div>
                          <input
                            type="text"
                            value={slide.title}
                            onChange={(e) => {
                              const slides = [...(settingsDraft.bannerSlides ?? [])];
                              slides[idx] = { ...slides[idx]!, title: e.target.value };
                              setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                            }}
                            className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-lg p-2 text-white text-xs focus:border-[#ffd025]"
                            placeholder="Título del slide"
                          />
                          <textarea
                            rows={2}
                            value={slide.description}
                            onChange={(e) => {
                              const slides = [...(settingsDraft.bannerSlides ?? [])];
                              slides[idx] = { ...slides[idx]!, description: e.target.value };
                              setSettingsDraft({ ...settingsDraft, bannerSlides: slides });
                            }}
                            className="w-full bg-[#141414] border border-[#ffd025]/20 rounded-lg p-2 text-white text-xs focus:border-[#ffd025]"
                            placeholder="Descripción del slide"
                          />
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="pt-4 border-t border-gray-800 space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      🎯 Barra de Avisos (Ticker)
                    </h3>
                    <p className="text-[10px] text-gray-500">Textos que aparecen en la barra animada arriba del header. Agrega hashtags y emojis para más impacto.</p>
                    <div className="space-y-2">
                      {((settingsDraft.announcements ?? []) as string[]).map((txt, idx) => (
                        <div key={idx} className="flex gap-2 items-start">
                          <input
                            type="text"
                            value={txt}
                            onChange={(e) => {
                              const arr = [...((settingsDraft.announcements ?? []) as string[])];
                              arr[idx] = e.target.value;
                              setSettingsDraft({ ...settingsDraft, announcements: arr });
                            }}
                            className="flex-1 bg-[#141414] border border-[#7c3aed]/30 rounded-xl p-2.5 text-white text-xs focus:border-[#7c3aed]"
                            placeholder="Ej: 🔥 Ofertas del finde #promo"
                          />
                          <button
                            onClick={async () => {
                              try {
                                const resp = await fetch(`${import.meta.env.BASE_URL}api/ai/improve-announcement`, {
                                  method: "POST",
                                  headers: { "Content-Type": "application/json" },
                                  body: JSON.stringify({ text: txt }),
                                });
                                const data = await resp.json();
                                if (data.text) {
                                  const arr = [...((settingsDraft.announcements ?? []) as string[])];
                                  arr[idx] = data.text;
                                  setSettingsDraft({ ...settingsDraft, announcements: arr });
                                }
                              } catch {}
                            }}
                            className="p-2.5 bg-[#7c3aed]/20 text-[#7c3aed] hover:bg-[#7c3aed] hover:text-white rounded-xl transition-colors border border-[#7c3aed]/30"
                            title="Mejorar con IA"
                          >
                            <Sparkles size={14} />
                          </button>
                          <button
                            onClick={() => {
                              const arr = [...((settingsDraft.announcements ?? []) as string[])];
                              arr.splice(idx, 1);
                              setSettingsDraft({ ...settingsDraft, announcements: arr });
                            }}
                            className="p-2.5 bg-red-500/10 text-red-400 hover:bg-red-500 hover:text-white rounded-xl transition-colors border border-red-500/20"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      ))}
                      <button
                        onClick={() => {
                          const arr = [...((settingsDraft.announcements ?? []) as string[])];
                          arr.push("");
                          setSettingsDraft({ ...settingsDraft, announcements: arr });
                        }}
                        className="w-full py-2.5 bg-[#7c3aed]/10 text-[#7c3aed] hover:bg-[#7c3aed]/20 rounded-xl text-xs font-bold uppercase flex items-center justify-center gap-2 border border-[#7c3aed]/20 transition-colors"
                      >
                        <Plus size={14} /> Agregar aviso
                      </button>
                    </div>
                  </div>

                  <div className="pt-4 border-t border-gray-800 space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      📍 Contacto y Footer
                    </h3>
                    <div>
                      <label className="block text-xs font-bold text-[#ffd025] uppercase mb-2">
                        Teléfono / WhatsApp
                      </label>
                      <input
                        type="text"
                        value={settingsDraft.contactPhone ?? ""}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, contactPhone: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] font-mono text-sm"
                        placeholder="Ej: +56 9 1234 5678"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-[#ec4899] uppercase mb-2">
                        Dirección / Ubicación
                      </label>
                      <input
                        type="text"
                        value={settingsDraft.contactAddress ?? ""}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, contactAddress: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ec4899]/30 rounded-xl p-3 text-white focus:border-[#ec4899] text-sm"
                        placeholder="Ej: Av. Principal 123, Santiago"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-[#7c3aed] uppercase mb-2">
                        Horarios de Atención (texto informativo)
                      </label>
                      <textarea
                        value={settingsDraft.contactHours ?? ""}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, contactHours: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#7c3aed]/30 rounded-xl p-3 text-white focus:border-[#7c3aed] text-sm"
                        placeholder={"Lun-Vie: 10:00 - 22:00\nSáb-Dom: 11:00 - 23:00"}
                        rows={3}
                      />
                    </div>
                  </div>

                  {/* Business hours for order validation */}
                  <div className="pt-4 border-t border-gray-800 space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      🕐 Horario de Recepción de Pedidos
                    </h3>
                    <p className="text-[11px] text-gray-500">
                      Fuera de este rango, el botón de pedidos se desactiva y el servidor rechaza nuevas órdenes.
                    </p>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-amber-400 uppercase mb-2">
                          Apertura
                        </label>
                        <input
                          type="time"
                          value={settingsDraft.openTime ?? "11:00"}
                          onChange={(e) =>
                            setSettingsDraft({ ...settingsDraft, openTime: e.target.value })
                          }
                          className="w-full bg-[#141414] border border-amber-500/30 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-amber-400 uppercase mb-2">
                          Cierre
                        </label>
                        <input
                          type="time"
                          value={settingsDraft.closeTime ?? "23:00"}
                          onChange={(e) =>
                            setSettingsDraft({ ...settingsDraft, closeTime: e.target.value })
                          }
                          className="w-full bg-[#141414] border border-amber-500/30 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-amber-400 uppercase mb-2">
                        Mínimo de compra — Delivery ($)
                      </label>
                      <input
                        type="number"
                        min={0}
                        step={1000}
                        value={settingsDraft.deliveryMinimum ?? 10000}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, deliveryMinimum: Number(e.target.value) })
                        }
                        className="w-full bg-[#141414] border border-amber-500/30 rounded-xl p-3 text-white focus:border-amber-400 text-sm font-mono"
                        placeholder="10000"
                      />
                      <p className="text-[10px] text-gray-600 mt-1">
                        Los pedidos con Retiro en tienda no tienen mínimo.
                      </p>
                    </div>
                  </div>

                  <div className="pt-4 border-t border-gray-800 space-y-4">
                    <h3 className="text-sm font-bold text-gray-300 uppercase flex items-center gap-2 border-b border-gray-800 pb-2">
                      WhatsApp y Créditos
                    </h3>
                    <div>
                      <label className="block text-xs font-bold text-[#ffd025] uppercase mb-2">
                        Número WhatsApp (Pedidos)
                      </label>
                      <input
                        type="text"
                        value={settingsDraft.whatsapp}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, whatsapp: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] font-mono text-sm"
                        placeholder="Ej: 56911223344"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Nombre de la Agencia
                      </label>
                      <input
                        type="text"
                        value={settingsDraft.agencyName}
                        onChange={(e) =>
                          setSettingsDraft({ ...settingsDraft, agencyName: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                        placeholder="Ej: Web Studio"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Logo de la Agencia (Opcional)
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={settingsDraft.agencyLogo}
                          onChange={(e) =>
                            setSettingsDraft({ ...settingsDraft, agencyLogo: e.target.value })
                          }
                          onBlur={(e) =>
                            resolveImageUrl(e.target.value, (r) =>
                              setSettingsDraft((p) => ({ ...p, agencyLogo: r })),
                            )
                          }
                          className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white focus:border-[#ffd025] text-sm"
                          placeholder="URL del logo"
                        />
                        <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                          <Upload size={20} />
                          <input
                            type="file"
                            accept="image/*"
                            className="hidden"
                            onChange={(e) =>
                              handleImageUpload(e, (url) =>
                                setSettingsDraft({ ...settingsDraft, agencyLogo: url }),
                              )
                            }
                          />
                        </label>
                        {settingsDraft.agencyLogo.includes("/storage/objects/") && (
                          <button
                            type="button"
                            title="Eliminar imagen del servidor"
                            onClick={() =>
                              handleDeleteStorageImage(settingsDraft.agencyLogo, () =>
                                setSettingsDraft((p) => ({ ...p, agencyLogo: "" })),
                              )
                            }
                            className="bg-red-500/10 text-red-400 px-4 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30"
                          >
                            <Trash2 size={20} />
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <button
                    onClick={saveSettings}
                    className="w-full py-4 mt-4 bg-[#ffd025] text-[#141414] rounded-xl font-black uppercase hover:bg-[#e5b81a] transition-all shadow-lg"
                  >
                    Guardar Ajustes a la Nube
                  </button>
                </div>
              </div>
            )}

            {adminTab === "settings" && (
              <DeliveryLocationsAdminPanel />
            )}

            {adminTab === "classifications" && (
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 animate-fade-in">
                <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
                  <h2 className="text-xl font-black uppercase mb-2 flex items-center gap-2">
                    <Layers size={20} className="text-[#ffd025]" /> Categorías
                  </h2>
                  <p className="text-xs text-gray-400 mb-6">
                    Pestañas superiores. Arrastra para reordenar.
                  </p>
                  <div className="flex gap-2 mb-6">
                    <input
                      type="text"
                      value={newCategoryName}
                      onChange={(e) => setNewCategoryName(e.target.value)}
                      placeholder="Nueva categoría..."
                      onKeyDown={(e) => e.key === "Enter" && handleAddCategory()}
                      className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                    />
                    <button
                      onClick={handleAddCategory}
                      className="px-4 py-3 bg-[#ffd025]/10 text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] rounded-xl font-bold transition-colors"
                    >
                      <Plus size={20} />
                    </button>
                  </div>
                  <div className="space-y-2">
                    {categoriesData.map((cat, index) => (
                      <div
                        key={cat.id}
                        draggable
                        onDragStart={(e) => handleDragStart(e, index)}
                        onDragOver={(e) => e.preventDefault()}
                        onDrop={(e) => handleDrop(e, index)}
                        className="flex justify-between items-center bg-[#141414] p-3 rounded-xl border border-gray-800 cursor-move hover:border-[#ffd025]/50 transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <GripVertical size={16} className="text-gray-500" />
                          <span className="font-bold text-white text-sm">{cat.name}</span>
                        </div>
                        <button
                          onClick={() => deleteCategoryHandler(cat.id)}
                          className="text-red-500 hover:text-red-400 p-1"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    ))}
                  </div>
                  <button
                    onClick={() =>
                      reorderCategoriesMut.mutate(
                        { data: { ids: categoriesData.map((c) => c.id) } },
                        {
                          onSuccess: () => {
                            setSavedCats(true);
                            setTimeout(() => setSavedCats(false), 2500);
                          },
                        },
                      )
                    }
                    disabled={reorderCategoriesMut.isPending}
                    className="mt-4 w-full py-3 rounded-xl font-black uppercase text-sm transition-all flex items-center justify-center gap-2 bg-[#ffd025] text-[#141414] hover:bg-[#e5b81a] disabled:opacity-60"
                  >
                    {savedCats ? (
                      <><CheckCircle size={16} /> Guardado</>
                    ) : reorderCategoriesMut.isPending ? (
                      "Guardando..."
                    ) : (
                      "Guardar Categorías"
                    )}
                  </button>
                </div>

                <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
                  <h2 className="text-xl font-black uppercase mb-2 flex items-center gap-2">
                    <Tag size={20} className="text-[#ffd025]" /> Pasillos
                  </h2>
                  <p className="text-xs text-gray-400 mb-6">
                    Títulos que separan productos en la lista.
                  </p>
                  <div className="flex gap-2 mb-6">
                    <input
                      type="text"
                      value={newAisleName}
                      onChange={(e) => setNewAisleName(e.target.value)}
                      placeholder="Nuevo pasillo..."
                      onKeyDown={(e) => e.key === "Enter" && handleAddAisle()}
                      className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                    />
                    <button
                      onClick={handleAddAisle}
                      className="px-4 py-3 bg-[#ffd025]/10 text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] rounded-xl font-bold transition-colors"
                    >
                      <Plus size={20} />
                    </button>
                  </div>
                  <div className="space-y-2 max-h-[32rem] overflow-y-auto pr-2">
                    {aislesData.map((aisle, aisleIndex) => {
                      const titleDraft = aisleBannerDrafts[`t_${aisle.id}`] ?? aisle.bannerTitle ?? "";
                      const subDraft = aisleBannerDrafts[`s_${aisle.id}`] ?? aisle.bannerSubtitle ?? "";
                      const textDirty =
                        titleDraft !== (aisle.bannerTitle ?? "") ||
                        subDraft !== (aisle.bannerSubtitle ?? "");
                      const saveBanner = (
                        partial: { bannerImage?: string; bannerTitle?: string; bannerSubtitle?: string },
                        msg?: string,
                      ) => {
                        const newData = {
                          bannerImage: partial.bannerImage ?? aisle.bannerImage ?? "",
                          bannerTitle: partial.bannerTitle ?? aisle.bannerTitle ?? "",
                          bannerSubtitle: partial.bannerSubtitle ?? aisle.bannerSubtitle ?? "",
                        };
                        queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                          old
                            ? { ...old, aisles: old.aisles.map((a: any) => (a.id === aisle.id ? { ...a, ...newData } : a)) }
                            : old,
                        );
                        if (msg) showToast(msg);
                        updateAisleMut.mutate(
                          { id: aisle.id, data: newData },
                          { onError: refreshMenu },
                        );
                      };
                      return (
                        <div
                          key={aisle.id}
                          className="bg-[#141414] p-3 rounded-xl border border-gray-800 space-y-2"
                        >
                          <div className="flex justify-between items-center">
                            <div className="flex items-center gap-1">
                              <div className="flex flex-col gap-0.5">
                                <button
                                  onClick={() => moveAisle(aisleIndex, -1)}
                                  disabled={aisleIndex === 0}
                                  className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                >
                                  <ChevronUp size={14} />
                                </button>
                                <button
                                  onClick={() => moveAisle(aisleIndex, 1)}
                                  disabled={aisleIndex === aislesData.length - 1}
                                  className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                                >
                                  <ChevronDown size={14} />
                                </button>
                              </div>
                              <span className="font-bold text-white text-sm pl-1">{aisle.name}</span>
                            </div>
                            <button
                              onClick={() => deleteAisleHandler(aisle.id)}
                              className="text-red-500 hover:text-red-400 p-1"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                          {aisle.bannerImage && (
                            <img
                              src={aisle.bannerImage}
                              alt="banner"
                              className="w-full h-12 object-cover rounded-lg border border-gray-800"
                            />
                          )}
                          <div className="space-y-1.5">
                            <div className="flex gap-2">
                              <input
                                type="text"
                                value={aisle.bannerImage || ""}
                                onChange={(e) =>
                                  saveBanner({ bannerImage: e.target.value })
                                }
                                onBlur={(e) =>
                                  resolveImageUrl(e.target.value, (r) =>
                                    saveBanner({ bannerImage: r }, "Banner actualizado"),
                                  )
                                }
                                placeholder="URL imagen banner"
                                className="flex-1 bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                              />
                              <label className="cursor-pointer px-3 py-2 bg-[#ffd025]/10 text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] rounded-lg text-xs font-black text-center transition flex items-center">
                                <Upload size={14} />
                                <input
                                  type="file"
                                  accept="image/*"
                                  className="hidden"
                                  onChange={(e) =>
                                    handleImageUpload(
                                      e,
                                      (url) => saveBanner({ bannerImage: url }, "Banner actualizado"),
                                      { skipCompression: true },
                                    )
                                  }
                                />
                              </label>
                              {aisle.bannerImage && (
                                <button
                                  onClick={() => {
                                    if (aisle.bannerImage?.includes("/storage/objects/")) {
                                      handleDeleteStorageImage(aisle.bannerImage, () =>
                                        saveBanner({ bannerImage: "" }, "Imagen eliminada"),
                                      );
                                    } else {
                                      saveBanner({ bannerImage: "" }, "Imagen eliminada");
                                    }
                                  }}
                                  className="px-3 py-2 bg-red-500/10 text-red-400 hover:bg-red-500 hover:text-white rounded-lg text-xs font-black transition"
                                >
                                  <Trash2 size={14} />
                                </button>
                              )}
                            </div>
                          </div>
                          <input
                            type="text"
                            value={titleDraft}
                            onChange={(e) =>
                              setAisleBannerDrafts((p) => ({ ...p, [`t_${aisle.id}`]: e.target.value }))
                            }
                            placeholder="Título 🍻 #ofertas"
                            className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                          />
                          <input
                            type="text"
                            value={subDraft}
                            onChange={(e) =>
                              setAisleBannerDrafts((p) => ({ ...p, [`s_${aisle.id}`]: e.target.value }))
                            }
                            placeholder="Subtítulo 🔥 ¡aprovecha!"
                            className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                          />
                          <div className="flex gap-2">
                            <button
                              onClick={async () => {
                                if (!titleDraft.trim() && !subDraft.trim()) {
                                  showToast("Escribe un borrador primero");
                                  return;
                                }
                                try {
                                  showToast("Mejorando con IA...");
                                  const res = await fetch(
                                    `${import.meta.env.BASE_URL}api/ai/improve-banner`,
                                    {
                                      method: "POST",
                                      headers: { "Content-Type": "application/json" },
                                      body: JSON.stringify({
                                        title: titleDraft,
                                        subtitle: subDraft,
                                        context: aisle.name,
                                      }),
                                    },
                                  );
                                  if (!res.ok) throw new Error("ai failed");
                                  const data = await res.json();
                                  setAisleBannerDrafts((p) => ({
                                    ...p,
                                    [`t_${aisle.id}`]: data.title ?? titleDraft,
                                    [`s_${aisle.id}`]: data.subtitle ?? subDraft,
                                  }));
                                  showToast("✨ Texto mejorado");
                                } catch {
                                  showToast("Error al mejorar texto");
                                }
                              }}
                              className="flex-1 px-3 py-2 bg-gradient-to-r from-[#ff3b8a] to-[#9d4edd] text-white rounded-lg text-xs font-black hover:scale-[1.02] transition flex items-center justify-center gap-1"
                            >
                              ✨ Mejorar con IA
                            </button>
                            {textDirty && (
                              <button
                                onClick={() => {
                                  saveBanner(
                                    { bannerTitle: titleDraft, bannerSubtitle: subDraft },
                                    "Textos actualizados",
                                  );
                                  setAisleBannerDrafts((p) => {
                                    const n = { ...p };
                                    delete n[`t_${aisle.id}`];
                                    delete n[`s_${aisle.id}`];
                                    return n;
                                  });
                                }}
                                className="flex-1 px-3 py-2 bg-[#ffd025] text-[#141414] rounded-lg text-xs font-black hover:scale-[1.02] transition"
                              >
                                Guardar
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <button
                    onClick={() =>
                      reorderAislesMut.mutate(
                        { data: { ids: aislesData.map((a) => a.id) } },
                        {
                          onSuccess: () => {
                            setSavedAisles(true);
                            setTimeout(() => setSavedAisles(false), 2500);
                          },
                        },
                      )
                    }
                    disabled={reorderAislesMut.isPending}
                    className="mt-4 w-full py-3 rounded-xl font-black uppercase text-sm transition-all flex items-center justify-center gap-2 bg-[#ffd025] text-[#141414] hover:bg-[#e5b81a] disabled:opacity-60"
                  >
                    {savedAisles ? (
                      <><CheckCircle size={16} /> Guardado</>
                    ) : reorderAislesMut.isPending ? (
                      "Guardando..."
                    ) : (
                      "Guardar Pasillos"
                    )}
                  </button>
                </div>

                <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
                  <h2 className="text-xl font-black uppercase mb-2 flex items-center gap-2">
                    <Tag size={20} className="text-[#ffd025]" /> Subcategorías
                  </h2>
                  <p className="text-xs text-gray-400 mb-6">
                    Etiquetas opcionales para destacar en la imagen.
                  </p>
                  <div className="flex gap-2 mb-6">
                    <input
                      type="text"
                      value={newSubcatName}
                      onChange={(e) => setNewSubcatName(e.target.value)}
                      placeholder="Nueva subcategoría..."
                      onKeyDown={(e) => e.key === "Enter" && handleAddSubcat()}
                      className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                    />
                    <button
                      onClick={handleAddSubcat}
                      className="px-4 py-3 bg-[#ffd025]/10 text-[#ffd025] hover:bg-[#ffd025] hover:text-[#141414] rounded-xl font-bold transition-colors"
                    >
                      <Plus size={20} />
                    </button>
                  </div>
                  <div className="space-y-2 max-h-96 overflow-y-auto pr-2">
                    {subcategoriesData.map((sub, subIndex) => (
                      <div
                        key={sub.id}
                        className="flex justify-between items-center bg-[#141414] p-3 rounded-xl border border-gray-800"
                      >
                        <div className="flex items-center gap-1">
                          <div className="flex flex-col gap-0.5">
                            <button
                              onClick={() => moveSubcat(subIndex, -1)}
                              disabled={subIndex === 0}
                              className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                            >
                              <ChevronUp size={14} />
                            </button>
                            <button
                              onClick={() => moveSubcat(subIndex, 1)}
                              disabled={subIndex === subcategoriesData.length - 1}
                              className="p-0.5 text-gray-400 hover:text-[#ffd025] disabled:opacity-20 disabled:cursor-not-allowed transition-colors"
                            >
                              <ChevronDown size={14} />
                            </button>
                          </div>
                          <span className="font-bold text-white text-sm pl-1">{sub.name}</span>
                        </div>
                        <button
                          onClick={() => deleteSubcategoryHandler(sub.id)}
                          className="text-red-500 hover:text-red-400 p-1"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    ))}
                  </div>
                  <button
                    onClick={() =>
                      reorderSubcategoriesMut.mutate(
                        { data: { ids: subcategoriesData.map((s) => s.id) } },
                        {
                          onSuccess: () => {
                            setSavedSubcats(true);
                            setTimeout(() => setSavedSubcats(false), 2500);
                          },
                        },
                      )
                    }
                    disabled={reorderSubcategoriesMut.isPending}
                    className="mt-4 w-full py-3 rounded-xl font-black uppercase text-sm transition-all flex items-center justify-center gap-2 bg-[#ffd025] text-[#141414] hover:bg-[#e5b81a] disabled:opacity-60"
                  >
                    {savedSubcats ? (
                      <><CheckCircle size={16} /> Guardado</>
                    ) : reorderSubcategoriesMut.isPending ? (
                      "Guardando..."
                    ) : (
                      "Guardar Subcategorías"
                    )}
                  </button>
                </div>
              </div>
            )}

            {adminTab === "social" && <CommunityAdminPanel />}

            {adminTab === "products" && (() => {
              const productFormFields = (
                <>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="col-span-2">
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Nombre
                      </label>
                      <input
                        required
                        type="text"
                        value={formState.name}
                        onChange={(e) => setFormState({ ...formState, name: e.target.value })}
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                      />
                    </div>
                    <div className="col-span-2">
                      <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                        Precio ($)
                      </label>
                      <input
                        required
                        type="number"
                        min="0"
                        value={formState.price}
                        onChange={(e) =>
                          setFormState({ ...formState, price: e.target.value })
                        }
                        className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                      />
                    </div>
                  </div>

                  <div className="p-4 bg-[#141414] rounded-xl border border-gray-800 grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-[10px] font-bold text-[#ffd025] uppercase mb-1">
                        Categoría
                      </label>
                      <select
                        required
                        value={formState.category}
                        onChange={(e) =>
                          setFormState({ ...formState, category: e.target.value })
                        }
                        className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-[#ffd025]"
                      >
                        <option value="">Seleccione...</option>
                        {categories.map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-[#ffd025] uppercase mb-1">
                        Pasillo
                      </label>
                      <select
                        required
                        value={formState.aisle}
                        onChange={(e) =>
                          setFormState({ ...formState, aisle: e.target.value })
                        }
                        className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-[#ffd025]"
                      >
                        <option value="">Seleccione...</option>
                        {aisles.map((a) => (
                          <option key={a} value={a}>
                            {a}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-[#ffd025] uppercase mb-1">
                        Subcat. (Opc.)
                      </label>
                      <select
                        value={formState.subcategory}
                        onChange={(e) =>
                          setFormState({ ...formState, subcategory: e.target.value })
                        }
                        className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-[#ffd025]"
                      >
                        <option value="">Ninguna...</option>
                        {subcategories.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  <div className="space-y-3 p-4 bg-[#141414] rounded-xl border border-blue-900/50">
                    <div>
                      <label className="block text-[10px] font-bold text-blue-400 uppercase mb-1">
                        Opciones / Variantes (Opcional)
                      </label>
                      <input
                        type="text"
                        placeholder="Ej: Elige tu bebida"
                        value={formState.optionsTitle}
                        onChange={(e) =>
                          setFormState({ ...formState, optionsTitle: e.target.value })
                        }
                        className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm mb-2 focus:border-blue-500"
                      />
                      <textarea
                        rows={2}
                        placeholder="Ej: Coca Cola, Sprite, Fanta (Separadas por comas)"
                        value={formState.optionsString}
                        onChange={(e) =>
                          setFormState({ ...formState, optionsString: e.target.value })
                        }
                        className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-blue-500"
                      ></textarea>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-gray-400 uppercase mb-2">
                      Imagen URL
                    </label>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        value={formState.image}
                        onChange={(e) =>
                          setFormState({ ...formState, image: e.target.value })
                        }
                        onBlur={(e) =>
                          resolveImageUrl(e.target.value, (resolved) =>
                            setFormState((prev) => ({ ...prev, image: resolved })),
                          )
                        }
                        placeholder="URL o sube una imagen cuadrada"
                        className="flex-1 bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025]"
                      />
                      <label className="bg-[#ffd025]/10 text-[#ffd025] px-4 rounded-xl flex items-center justify-center cursor-pointer hover:bg-[#ffd025]/20 border border-[#ffd025]/30">
                        <Upload size={20} />
                        <input
                          type="file"
                          accept="image/*"
                          className="hidden"
                          onChange={(e) =>
                            handleImageUpload(e, (url) =>
                              setFormState({ ...formState, image: url }),
                            )
                          }
                        />
                      </label>
                      {formState.image.includes("/storage/objects/") && (
                        <button
                          type="button"
                          title="Eliminar imagen del servidor"
                          onClick={() =>
                            handleDeleteStorageImage(formState.image, () =>
                              setFormState({ ...formState, image: "" }),
                            )
                          }
                          className="bg-red-500/10 text-red-400 px-4 rounded-xl flex items-center justify-center hover:bg-red-500/20 border border-red-500/30"
                        >
                          <Trash2 size={20} />
                        </button>
                      )}
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => setFormState({ ...formState, oferta: !formState.oferta })}
                    className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border-2 transition-all font-bold text-sm ${
                      formState.oferta
                        ? "bg-gradient-to-r from-pink-600/30 to-yellow-500/20 border-pink-500 text-white"
                        : "bg-[#141414] border-gray-700 text-gray-400 hover:border-pink-500/50"
                    }`}
                  >
                    <span className="flex items-center gap-2">
                      🏷️ Marcar como OFERTA
                    </span>
                    <span className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                      formState.oferta ? "bg-pink-500 border-pink-500" : "border-gray-600"
                    }`}>
                      {formState.oferta && <span className="text-white text-[10px] font-black">✓</span>}
                    </span>
                  </button>

                  <div className="space-y-2">
                    <button
                      type="button"
                      onClick={() => setFormState({ ...formState, depositoEnabled: !formState.depositoEnabled })}
                      className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border-2 transition-all font-bold text-sm ${
                        formState.depositoEnabled
                          ? "bg-gradient-to-r from-cyan-600/30 to-blue-500/20 border-cyan-500 text-white"
                          : "bg-[#141414] border-gray-700 text-gray-400 hover:border-cyan-500/50"
                      }`}
                    >
                      <span className="flex items-center gap-2">
                        🍾 Habilitar Depósito Retornable
                      </span>
                      <span className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                        formState.depositoEnabled ? "bg-cyan-500 border-cyan-500" : "border-gray-600"
                      }`}>
                        {formState.depositoEnabled && <span className="text-white text-[10px] font-black">✓</span>}
                      </span>
                    </button>
                    {formState.depositoEnabled && (
                      <div className="flex items-center gap-2 px-1">
                        <label className="text-xs font-bold text-cyan-400 whitespace-nowrap">Monto depósito ($)</label>
                        <input
                          type="number"
                          min="0"
                          value={formState.depositoAmount}
                          onChange={(e) => setFormState({ ...formState, depositoAmount: e.target.value })}
                          className="flex-1 bg-[#141414] border border-cyan-500/50 rounded-lg px-3 py-1.5 text-white text-sm focus:border-cyan-400 focus:outline-none"
                          placeholder="500"
                        />
                      </div>
                    )}
                  </div>

                  <div className="space-y-2">
                    <button
                      type="button"
                      onClick={() => setFormState({ ...formState, transferenciaEnabled: !formState.transferenciaEnabled })}
                      className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border-2 transition-all font-bold text-sm ${
                        formState.transferenciaEnabled
                          ? "bg-gradient-to-r from-blue-600/30 to-indigo-500/20 border-blue-500 text-white"
                          : "bg-[#141414] border-gray-700 text-gray-400 hover:border-blue-500/50"
                      }`}
                    >
                      <span className="flex items-center gap-2">
                        💳 Cargo por Transferencia
                      </span>
                      <span className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                        formState.transferenciaEnabled ? "bg-blue-500 border-blue-500" : "border-gray-600"
                      }`}>
                        {formState.transferenciaEnabled && <span className="text-white text-[10px] font-black">✓</span>}
                      </span>
                    </button>
                    {formState.transferenciaEnabled && (
                      <div className="flex items-center gap-2 px-1">
                        <label className="text-xs font-bold text-blue-400 whitespace-nowrap">Monto cargo ($)</label>
                        <input
                          type="number"
                          min="0"
                          value={formState.transferenciaAmount}
                          onChange={(e) => setFormState({ ...formState, transferenciaAmount: e.target.value })}
                          className="flex-1 bg-[#141414] border border-blue-500/50 rounded-lg px-3 py-1.5 text-white text-sm focus:border-blue-400 focus:outline-none"
                          placeholder="0"
                        />
                      </div>
                    )}
                  </div>
                </>
              );

              return (
                <div className="grid grid-cols-1 lg:grid-cols-5 gap-8 animate-fade-in">
                  {editingProduct === null && (
                    <div className="lg:col-span-2">
                      <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20 sticky top-24">
                        <h2 className="text-xl font-black uppercase mb-6 flex items-center gap-2">
                          <Plus size={20} className="text-[#ffd025]" /> Añadir Producto
                        </h2>
                        <form onSubmit={saveProduct} className="space-y-4">
                          {productFormFields}
                          <div className="flex gap-2 pt-4">
                            <button
                              type="submit"
                              className="flex-1 py-3 bg-[#ffd025] text-[#141414] rounded-xl font-bold uppercase hover:bg-[#e5b81a] transition-colors shadow-lg"
                            >
                              Crear Producto
                            </button>
                          </div>
                        </form>
                      </div>
                    </div>
                  )}

                  <div className={editingProduct === null ? "lg:col-span-3" : "lg:col-span-5"}>
                    <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
                      <div className="flex items-center justify-between gap-4 mb-4 flex-wrap">
                        <div className="flex items-center gap-3">
                          <h2 className="text-xl font-black uppercase">
                            Catálogo Actual ({products.length})
                          </h2>
                          <button
                            type="button"
                            onClick={() => downloadProductsExcel(products)}
                            title="Descargar catálogo en Excel"
                            className="flex items-center gap-1.5 px-3 py-1.5 bg-green-700/20 border border-green-600/30 text-green-400 rounded-xl text-xs font-bold hover:bg-green-700/35 hover:border-green-500/50 transition-colors"
                          >
                            <Download size={13} />
                            Excel
                          </button>
                        </div>
                        <div className="relative flex-1 min-w-[180px] max-w-xs">
                          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none">
                            🔍
                          </span>
                          <input
                            type="text"
                            placeholder="Buscar producto..."
                            value={adminProductSearch}
                            onChange={(e) => setAdminProductSearch(e.target.value)}
                            className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl pl-9 pr-9 py-2 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder:text-gray-600"
                          />
                          {adminProductSearch && (
                            <button
                              type="button"
                              onClick={() => setAdminProductSearch("")}
                              className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-white transition-colors"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                      </div>
                      <div
                        ref={adminCategoriesRef}
                        className={`flex gap-2 overflow-x-auto scrollbar-hide pb-3 mb-4 -mx-2 px-2 border-b border-gray-800 scroll-fade-x scroll-smooth select-none transition-opacity ${adminProductSearch ? "opacity-30 pointer-events-none" : ""}`}
                      >
                      {["Todas", ...categories].map((cat) => {
                        const count =
                          cat === "Todas"
                            ? products.length
                            : products.filter((p) => p.category === cat).length;
                        const isActive = adminCategory === cat;
                        return (
                          <button
                            key={cat}
                            type="button"
                            onClick={() => setAdminCategory(cat)}
                            className={`flex items-center gap-2 px-4 py-2 rounded-full whitespace-nowrap text-xs font-bold uppercase transition-colors flex-shrink-0 ${
                              isActive
                                ? "bg-[#ffd025] text-[#141414] shadow-md"
                                : "bg-[#141414] text-gray-300 border border-gray-800 hover:border-[#ffd025]/40 hover:text-white"
                            }`}
                          >
                            <span>{cat}</span>
                            <span
                              className={`text-[10px] px-1.5 py-0.5 rounded-full font-bold ${
                                isActive
                                  ? "bg-[#141414]/20 text-[#141414]"
                                  : "bg-[#ffd025]/15 text-[#ffd025]"
                              }`}
                            >
                              {count}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                      <div className="space-y-2">
                        {(() => {
                          const normalize = (s: string) =>
                            s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
                          const filteredProducts = (() => {
                            const q = normalize(adminProductSearch.trim());
                            if (q.length > 0) {
                              const words = q.split(/\s+/).filter(Boolean);
                              return products.filter((p) => {
                                const haystack = normalize(
                                  [p.name, p.category, p.aisle, p.subcategory ?? ""].join(" ")
                                );
                                return words.every((w) => haystack.includes(w));
                              });
                            }
                            return adminCategory === "Todas"
                              ? products
                              : products.filter((p) => p.category === adminCategory);
                          })();
                          const groups = aisles.map((a) => ({
                            aisle: a,
                            items: filteredProducts.filter((p) => p.aisle === a),
                          }));
                          const ungrouped = filteredProducts.filter(
                            (p) => !aisles.includes(p.aisle),
                          );
                          if (ungrouped.length > 0) {
                            groups.push({ aisle: "Sin pasillo", items: ungrouped });
                          }
                          return groups
                            .filter((g) => g.items.length > 0)
                            .map(({ aisle, items }) => {
                              const collapsed = collapsedAisles[aisle];
                              return (
                              <div
                                key={aisle}
                                className="border border-gray-800 rounded-xl overflow-hidden bg-[#0f0f0f]"
                              >
                                <button
                                  type="button"
                                  onClick={() =>
                                    setCollapsedAisles({
                                      ...collapsedAisles,
                                      [aisle]: !collapsed,
                                    })
                                  }
                                  className="w-full flex items-center justify-between gap-2 p-3 bg-[#1a1a1a] hover:bg-[#222] transition-colors"
                                >
                                  <div className="flex items-center gap-2">
                                    {collapsed ? (
                                      <ChevronRight
                                        size={18}
                                        className="text-[#ffd025]"
                                      />
                                    ) : (
                                      <ChevronDown
                                        size={18}
                                        className="text-[#ffd025]"
                                      />
                                    )}
                                    <span className="font-bold uppercase text-sm text-white">
                                      {aisle}
                                    </span>
                                    <span className="text-[10px] bg-[#ffd025]/15 text-[#ffd025] px-2 py-0.5 rounded-full font-bold">
                                      {items.length}
                                    </span>
                                  </div>
                                </button>
                                {!collapsed && (
                                  <div className="space-y-3 p-2">
                                    {(() => {
                                      const subMap = new Map<string, typeof items>();
                                      const noSub: typeof items = [];
                                      for (const p of items) {
                                        const s = p.subcategory?.trim() || "";
                                        if (s) {
                                          if (!subMap.has(s)) subMap.set(s, []);
                                          subMap.get(s)!.push(p);
                                        } else {
                                          noSub.push(p);
                                        }
                                      }
                                      const renderProduct = (product: typeof items[0], groupIds: number[]) =>
                                        editingProduct === product.id ? (
                          <div
                            key={product.id}
                            className="bg-[#141414] p-4 rounded-xl border-2 border-[#ffd025] shadow-lg shadow-[#ffd025]/10 animate-fade-in"
                          >
                            <div className="flex items-center justify-between mb-3 pb-3 border-b border-[#ffd025]/20">
                              <h4 className="font-black uppercase text-sm flex items-center gap-2 text-[#ffd025]">
                                <Edit3 size={16} /> Editando producto
                              </h4>
                              <button
                                type="button"
                                onClick={cancelEditing}
                                className="p-1.5 text-gray-400 hover:text-white hover:bg-white/10 rounded-lg transition-colors"
                                aria-label="Cerrar"
                              >
                                <X size={18} />
                              </button>
                            </div>
                            <form onSubmit={saveProduct} className="space-y-3">
                              <div className="flex gap-3 items-start">
                                <div className="relative flex-shrink-0">
                                  {formState.image ? (
                                    <img
                                      src={formState.image}
                                      alt=""
                                      className="w-16 h-16 rounded-lg object-contain bg-[#1a1a1a] border border-gray-700"
                                      onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = "none"; }}
                                    />
                                  ) : (
                                    <div className="w-16 h-16 rounded-lg bg-[#1a1a1a] border border-gray-700" />
                                  )}
                                  <label className="absolute -bottom-1 -right-1 bg-[#ffd025] text-[#141414] p-1 rounded-full cursor-pointer hover:bg-[#e5b81a] shadow-md">
                                    <Upload size={12} />
                                    <input
                                      type="file"
                                      accept="image/*"
                                      className="hidden"
                                      onChange={(e) =>
                                        handleImageUpload(e, (url) =>
                                          setFormState({ ...formState, image: url }),
                                        )
                                      }
                                    />
                                  </label>
                                  {formState.image.includes("/storage/objects/") && (
                                    <button
                                      type="button"
                                      title="Eliminar imagen del servidor"
                                      onClick={() =>
                                        handleDeleteStorageImage(formState.image, () =>
                                          setFormState({ ...formState, image: "" }),
                                        )
                                      }
                                      className="absolute -top-1 -right-1 bg-red-500 text-white p-1 rounded-full hover:bg-red-600 shadow-md"
                                    >
                                      <Trash2 size={12} />
                                    </button>
                                  )}
                                </div>
                                <div className="flex-1 flex flex-col gap-2">
                                  <input
                                    type="text"
                                    placeholder="Pegar link de imagen (https://...)"
                                    value={formState.image.includes("/storage/objects/") ? "" : formState.image}
                                    onChange={(e) =>
                                      setFormState({ ...formState, image: e.target.value })
                                    }
                                    onBlur={(e) =>
                                      resolveImageUrl(e.target.value, (resolved) =>
                                        setFormState((prev) => ({ ...prev, image: resolved })),
                                      )
                                    }
                                    className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none placeholder-gray-600"
                                  />
                                  <div className="grid grid-cols-3 gap-2">
                                  <input
                                    required
                                    type="text"
                                    placeholder="Nombre"
                                    value={formState.name}
                                    onChange={(e) =>
                                      setFormState({ ...formState, name: e.target.value })
                                    }
                                    className="col-span-2 bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-[#ffd025] focus:outline-none"
                                  />
                                  <input
                                    required
                                    type="number"
                                    min="0"
                                    placeholder="Precio"
                                    value={formState.price}
                                    onChange={(e) =>
                                      setFormState({ ...formState, price: e.target.value })
                                    }
                                    className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-sm focus:border-[#ffd025] focus:outline-none"
                                  />
                                </div>
                                </div>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                                <select
                                  required
                                  value={formState.category}
                                  onChange={(e) =>
                                    setFormState({ ...formState, category: e.target.value })
                                  }
                                  className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                >
                                  <option value="">Categoría...</option>
                                  {categories.map((c) => (
                                    <option key={c} value={c}>
                                      {c}
                                    </option>
                                  ))}
                                </select>
                                <select
                                  required
                                  value={formState.aisle}
                                  onChange={(e) =>
                                    setFormState({ ...formState, aisle: e.target.value })
                                  }
                                  className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                >
                                  <option value="">Pasillo...</option>
                                  {aisles.map((a) => (
                                    <option key={a} value={a}>
                                      {a}
                                    </option>
                                  ))}
                                </select>
                                <select
                                  value={formState.subcategory}
                                  onChange={(e) =>
                                    setFormState({ ...formState, subcategory: e.target.value })
                                  }
                                  className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                                >
                                  <option value="">Subcat. (opc)...</option>
                                  {subcategories.map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                              </div>

                              <button
                                type="button"
                                onClick={() => setFormState({ ...formState, oferta: !formState.oferta })}
                                className={`w-full flex items-center justify-between px-3 py-2 rounded-lg border-2 transition-all text-xs font-bold ${
                                  formState.oferta
                                    ? "bg-gradient-to-r from-pink-600/30 to-yellow-500/20 border-pink-500 text-white"
                                    : "bg-[#141414] border-gray-700 text-gray-400 hover:border-pink-500/50"
                                }`}
                              >
                                <span>🏷️ {formState.oferta ? "¡EN OFERTA! 🔥" : "Marcar como Oferta"}</span>
                                <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${
                                  formState.oferta ? "bg-pink-500 border-pink-500" : "border-gray-600"
                                }`}>
                                  {formState.oferta && <span className="text-white text-[8px]">✓</span>}
                                </span>
                              </button>

                              <div className="space-y-1.5">
                                <button
                                  type="button"
                                  onClick={() => setFormState({ ...formState, depositoEnabled: !formState.depositoEnabled })}
                                  className={`w-full flex items-center justify-between px-3 py-2 rounded-lg border-2 transition-all text-xs font-bold ${
                                    formState.depositoEnabled
                                      ? "bg-gradient-to-r from-cyan-600/30 to-blue-500/20 border-cyan-500 text-white"
                                      : "bg-[#141414] border-gray-700 text-gray-400 hover:border-cyan-500/50"
                                  }`}
                                >
                                  <span>🍾 {formState.depositoEnabled ? "Depósito activo" : "Habilitar Depósito"}</span>
                                  <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${
                                    formState.depositoEnabled ? "bg-cyan-500 border-cyan-500" : "border-gray-600"
                                  }`}>
                                    {formState.depositoEnabled && <span className="text-white text-[8px]">✓</span>}
                                  </span>
                                </button>
                                {formState.depositoEnabled && (
                                  <div className="flex items-center gap-2">
                                    <label className="text-[10px] font-bold text-cyan-400 whitespace-nowrap">Monto ($)</label>
                                    <input
                                      type="number"
                                      min="0"
                                      value={formState.depositoAmount}
                                      onChange={(e) => setFormState({ ...formState, depositoAmount: e.target.value })}
                                      className="flex-1 bg-[#1a1a1a] border border-cyan-500/50 rounded-lg px-2 py-1 text-white text-xs focus:border-cyan-400 focus:outline-none"
                                      placeholder="500"
                                    />
                                  </div>
                                )}
                              </div>

                              <div className="space-y-1.5">
                                <button
                                  type="button"
                                  onClick={() => setFormState({ ...formState, transferenciaEnabled: !formState.transferenciaEnabled })}
                                  className={`w-full flex items-center justify-between px-3 py-2 rounded-lg border-2 transition-all text-xs font-bold ${
                                    formState.transferenciaEnabled
                                      ? "bg-gradient-to-r from-blue-600/30 to-indigo-500/20 border-blue-500 text-white"
                                      : "bg-[#141414] border-gray-700 text-gray-400 hover:border-blue-500/50"
                                  }`}
                                >
                                  <span>💳 {formState.transferenciaEnabled ? "Cargo transferencia activo" : "Habilitar Cargo Transferencia"}</span>
                                  <span className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${
                                    formState.transferenciaEnabled ? "bg-blue-500 border-blue-500" : "border-gray-600"
                                  }`}>
                                    {formState.transferenciaEnabled && <span className="text-white text-[8px]">✓</span>}
                                  </span>
                                </button>
                                {formState.transferenciaEnabled && (
                                  <div className="flex items-center gap-2">
                                    <label className="text-[10px] font-bold text-blue-400 whitespace-nowrap">Monto ($)</label>
                                    <input
                                      type="number"
                                      min="0"
                                      value={formState.transferenciaAmount}
                                      onChange={(e) => setFormState({ ...formState, transferenciaAmount: e.target.value })}
                                      className="flex-1 bg-[#1a1a1a] border border-blue-500/50 rounded-lg px-2 py-1 text-white text-xs focus:border-blue-400 focus:outline-none"
                                      placeholder="0"
                                    />
                                  </div>
                                )}
                              </div>

                              <input
                                type="text"
                                placeholder="URL de la imagen"
                                value={formState.image}
                                onChange={(e) =>
                                  setFormState({ ...formState, image: e.target.value })
                                }
                                onBlur={(e) =>
                                  resolveImageUrl(e.target.value, (resolved) =>
                                    setFormState((prev) => ({ ...prev, image: resolved })),
                                  )
                                }
                                className="w-full bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-[#ffd025] focus:outline-none"
                              />

                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                <input
                                  type="text"
                                  placeholder="Variantes: título (opc)"
                                  value={formState.optionsTitle}
                                  onChange={(e) =>
                                    setFormState({
                                      ...formState,
                                      optionsTitle: e.target.value,
                                    })
                                  }
                                  className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-blue-500 focus:outline-none"
                                />
                                <input
                                  type="text"
                                  placeholder="Variantes (separadas por coma)"
                                  value={formState.optionsString}
                                  onChange={(e) =>
                                    setFormState({
                                      ...formState,
                                      optionsString: e.target.value,
                                    })
                                  }
                                  className="bg-[#1a1a1a] border border-gray-700 rounded-lg p-2 text-white text-xs focus:border-blue-500 focus:outline-none"
                                />
                              </div>

                              <div className="flex gap-2 pt-1">
                                <button
                                  type="button"
                                  onClick={cancelEditing}
                                  className="px-4 py-2 bg-gray-800 text-white rounded-lg font-bold uppercase text-xs hover:bg-gray-700 transition-colors"
                                >
                                  Cancelar
                                </button>
                                <button
                                  type="submit"
                                  className="flex-1 py-2 bg-[#ffd025] text-[#141414] rounded-lg font-bold uppercase text-xs hover:bg-[#e5b81a] transition-colors shadow"
                                >
                                  Guardar Cambios
                                </button>
                              </div>
                            </form>
                          </div>
                        ) : (
                          <div
                            key={product.id}
                            className="flex flex-col sm:flex-row gap-4 items-center bg-[#141414] p-3 rounded-xl border border-gray-800 hover:border-gray-600 transition-colors"
                          >
                            <div className="relative flex-shrink-0">
                              {product.image ? (
                                <img
                                  loading="lazy"
                                  src={product.image}
                                  alt=""
                                  className="w-16 h-16 rounded-lg object-contain bg-[#1a1a1a]"
                                  onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = "none"; }}
                                />
                              ) : (
                                <div className="w-16 h-16 rounded-lg bg-[#1a1a1a] border border-gray-800" />
                              )}
                              {product.image && (
                                <span
                                  title={product.image.includes("/storage/objects/") ? "Imagen guardada en el servidor" : "Imagen via link externo"}
                                  className={`absolute -bottom-1 -right-1 p-1 rounded-full shadow ${product.image.includes("/storage/objects/") ? "bg-emerald-600 text-white" : "bg-blue-600 text-white"}`}
                                >
                                  {product.image.includes("/storage/objects/") ? <HardDriveUpload size={10} /> : <Link2 size={10} />}
                                </span>
                              )}
                            </div>
                            <div className="flex-1 w-full text-center sm:text-left">
                              <h4 className="font-bold text-white text-sm md:text-base">
                                {product.name}
                              </h4>
                              <div className="flex flex-wrap justify-center sm:justify-start gap-1.5 mt-1 text-[10px] md:text-xs">
                                <span className="bg-gray-800 text-gray-300 px-2 py-0.5 rounded">
                                  {product.category}
                                </span>
                                <span className="bg-gray-800 text-gray-300 px-2 py-0.5 rounded">
                                  {product.aisle}
                                </span>
                                {product.subcategory && (
                                  <span className="bg-gray-700 text-[#ffd025] px-2 py-0.5 rounded border border-[#ffd025]/30">
                                    {product.subcategory}
                                  </span>
                                )}
                                <span className="text-[#ffd025] font-bold py-0.5">
                                  ${product.price.toLocaleString("es-CL")}
                                </span>
                              </div>
                              {product.options && product.options.length > 0 && (
                                <p className="text-[10px] text-blue-400 mt-1 line-clamp-1">
                                  V: {product.options.join(", ")}
                                </p>
                              )}
                            </div>
                            <div className="flex gap-1 w-full sm:w-auto justify-center items-center">
                              <div className="flex flex-col gap-0.5 mr-1">
                                <button
                                  title="Subir"
                                  onClick={() => moveProduct(product.id, "up", groupIds)}
                                  disabled={groupIds.indexOf(product.id) === 0}
                                  className="p-1.5 text-[#ffd025]/70 hover:text-[#ffd025] hover:bg-[#ffd025]/10 rounded-lg transition-colors disabled:opacity-20 disabled:cursor-not-allowed"
                                >
                                  <ChevronUp size={14} />
                                </button>
                                <button
                                  title="Bajar"
                                  onClick={() => moveProduct(product.id, "down", groupIds)}
                                  disabled={groupIds.indexOf(product.id) === groupIds.length - 1}
                                  className="p-1.5 text-[#ffd025]/70 hover:text-[#ffd025] hover:bg-[#ffd025]/10 rounded-lg transition-colors disabled:opacity-20 disabled:cursor-not-allowed"
                                >
                                  <ChevronDown size={14} />
                                </button>
                              </div>
                              <button
                                title={product.publishedSocial ? "Publicado en redes ✓ — Click para desmarcar" : "Marcar como publicado en redes sociales"}
                                onClick={() => {
                                  const newVal = !product.publishedSocial;
                                  queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                    old
                                      ? { ...old, products: old.products.map((p: any) => p.id === product.id ? { ...p, publishedSocial: newVal } : p) }
                                      : old,
                                  );
                                  updateProductMut.mutate(
                                    {
                                      id: product.id,
                                      data: {
                                        name: product.name,
                                        price: product.price,
                                        image: product.image,
                                        category: product.category,
                                        aisle: product.aisle,
                                        subcategory: product.subcategory ?? "",
                                        optionsTitle: product.optionsTitle ?? "",
                                        options: product.options ?? [],
                                        oferta: product.oferta ?? false,
                                        depositoEnabled: product.depositoEnabled ?? false,
                                        depositoAmount: product.depositoAmount ?? 500,
                                        publishedSocial: newVal,
                                        transferenciaEnabled: product.transferenciaEnabled ?? false,
                                        transferenciaAmount: product.transferenciaAmount ?? 0,
                                      },
                                    },
                                    { onError: refreshMenu },
                                  );
                                }}
                                className={`p-3 rounded-xl transition-colors ${
                                  product.publishedSocial
                                    ? "text-green-400 bg-green-500/15 hover:bg-green-500/25"
                                    : "text-gray-600 hover:text-green-400 hover:bg-green-500/10"
                                }`}
                              >
                                <CheckCircle size={18} />
                              </button>
                              <button
                                title={product.hidden ? "Oculto — click para mostrar" : "Visible — click para ocultar"}
                                onClick={() => {
                                  const newVal = !product.hidden;
                                  queryClient.setQueryData(getGetMenuQueryKey(), (old: any) =>
                                    old
                                      ? { ...old, products: old.products.map((p: any) => p.id === product.id ? { ...p, hidden: newVal } : p) }
                                      : old,
                                  );
                                  updateProductMut.mutate(
                                    {
                                      id: product.id,
                                      data: {
                                        name: product.name,
                                        price: product.price,
                                        image: product.image,
                                        category: product.category,
                                        aisle: product.aisle,
                                        subcategory: product.subcategory ?? "",
                                        optionsTitle: product.optionsTitle ?? "",
                                        options: product.options ?? [],
                                        oferta: product.oferta ?? false,
                                        depositoEnabled: product.depositoEnabled ?? false,
                                        depositoAmount: product.depositoAmount ?? 500,
                                        hidden: newVal,
                                        transferenciaEnabled: product.transferenciaEnabled ?? false,
                                        transferenciaAmount: product.transferenciaAmount ?? 0,
                                      },
                                    },
                                    { onError: refreshMenu },
                                  );
                                }}
                                className={`p-3 rounded-xl transition-colors ${
                                  product.hidden
                                    ? "text-amber-400 bg-amber-500/15 hover:bg-amber-500/25"
                                    : "text-gray-600 hover:text-amber-400 hover:bg-amber-500/10"
                                }`}
                              >
                                {product.hidden ? <EyeOff size={18} /> : <Eye size={18} />}
                              </button>
                              <button
                                onClick={() => startEditing(product)}
                                className="p-3 text-blue-400 hover:bg-blue-400/10 rounded-xl transition-colors"
                              >
                                <Edit3 size={18} />
                              </button>
                              <button
                                onClick={() => deleteProductHandler(product.id)}
                                className="p-3 text-red-500 hover:bg-red-500/10 rounded-xl transition-colors"
                              >
                                <Trash2 size={18} />
                              </button>
                            </div>
                          </div>
                        );
                                      return (
                                        <>
                                          {noSub.length > 0 && (
                                            <div className="space-y-2">
                                              {noSub.map((p) => renderProduct(p, noSub.map((q) => q.id)))}
                                            </div>
                                          )}
                                          {Array.from(subMap.entries()).map(([sub, subItems]) => (
                                            <div key={sub}>
                                              <div className="flex items-center gap-2 px-1 py-1 mb-1">
                                                <span className="text-[10px] font-bold uppercase text-[#ffd025]/70 tracking-wider">{sub}</span>
                                                <div className="flex-1 h-px bg-[#ffd025]/15" />
                                                <span className="text-[9px] text-gray-600">{subItems.length}</span>
                                              </div>
                                              <div className="space-y-2">
                                                {subItems.map((p) => renderProduct(p, subItems.map((q) => q.id)))}
                                              </div>
                                            </div>
                                          ))}
                                        </>
                                      );
                                    })()}
                                  </div>
                                )}
                              </div>
                            );
                          });
                      })()}
                      {products.length === 0 && (
                        <p className="text-gray-500 text-center py-8">
                          No hay productos en el catálogo.
                        </p>
                      )}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })()}

          </div>
        </div>
      )}

      <style
        dangerouslySetInnerHTML={{
          __html: `
        .scrollbar-hide::-webkit-scrollbar { display: none; }
        .scrollbar-hide { -ms-overflow-style: none; scrollbar-width: none; }
        .scroll-fade-x {
          -webkit-mask-image: linear-gradient(to right, transparent 0, black 28px, black calc(100% - 28px), transparent 100%);
          mask-image: linear-gradient(to right, transparent 0, black 28px, black calc(100% - 28px), transparent 100%);
        }
        @keyframes slideInRight { from { transform: translateX(100%); } to { transform: translateX(0); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .animate-slide-in-right { animation: slideInRight 0.3s cubic-bezier(0.16, 1, 0.3, 1); }
        .animate-fade-in { animation: fadeIn 0.4s ease-out forwards; }
      `,
        }}
      />
    </div>
  );
}

const ORDER_STATUSES = [
  {
    value: "Confirmado y en Preparación",
    label: "Confirmado y en Preparación",
    color: "bg-blue-500/15 text-blue-300 border-blue-500/40",
    activeColor: "bg-blue-500 text-white border-blue-500",
    whatsapp: true,
  },
  {
    value: "Listo para Retirar",
    label: "Listo para Retirar",
    color: "bg-amber-500/15 text-amber-300 border-amber-500/40",
    activeColor: "bg-amber-500 text-[#141414] border-amber-500",
    whatsapp: true,
  },
  {
    value: "Delivery en Camino",
    label: "Delivery en Camino",
    color: "bg-emerald-500/15 text-emerald-300 border-emerald-500/40",
    activeColor: "bg-emerald-500 text-[#141414] border-emerald-500",
    whatsapp: true,
  },
  {
    value: "Entregado",
    label: "✓ Entregado",
    color: "bg-green-700/20 text-green-400 border-green-600/40",
    activeColor: "bg-green-600 text-white border-green-600",
    whatsapp: false,
  },
] as const;

function DeliveryLocationsAdminPanel() {
  const queryClient = useQueryClient();
  const { data: locations = [] } = useListDeliveryLocations();
  const createMut = useCreateDeliveryLocation();
  const updateMut = useUpdateDeliveryLocation();
  const deleteMut = useDeleteDeliveryLocation();

  const [form, setForm] = useState({ name: "", price: "" });
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editForm, setEditForm] = useState({ name: "", price: "" });

  const invalidate = () => queryClient.invalidateQueries({ queryKey: getListDeliveryLocationsQueryKey() });

  const handleCreate = () => {
    const price = parseInt(form.price, 10);
    if (!form.name.trim() || isNaN(price)) return;
    createMut.mutate({ data: { name: form.name.trim(), price } }, { onSuccess: () => { setForm({ name: "", price: "" }); invalidate(); } });
  };

  const startEdit = (loc: DeliveryLocation) => {
    setEditingId(loc.id);
    setEditForm({ name: loc.name, price: String(loc.price) });
  };

  const handleUpdate = (id: number) => {
    const price = parseInt(editForm.price, 10);
    if (!editForm.name.trim() || isNaN(price)) return;
    updateMut.mutate({ id, data: { name: editForm.name.trim(), price } }, { onSuccess: () => { setEditingId(null); invalidate(); } });
  };

  const toggleActive = (loc: DeliveryLocation) => {
    updateMut.mutate({ id: loc.id, data: { active: !loc.active } }, { onSuccess: invalidate });
  };

  const handleDelete = (id: number) => {
    deleteMut.mutate({ id }, { onSuccess: invalidate });
  };

  return (
    <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20 max-w-2xl animate-fade-in">
      <h2 className="text-xl font-black uppercase mb-6 flex items-center gap-2">
        <span className="text-[#ffd025]">📍</span> Zonas de Delivery
      </h2>
      <div className="space-y-3 mb-6">
        {locations.length === 0 && (
          <p className="text-gray-500 text-sm">No hay zonas configuradas. Agrega una abajo.</p>
        )}
        {locations.map(loc => (
          <div key={loc.id} className="flex items-center gap-3 bg-[#141414] rounded-xl p-3 border border-gray-800">
            {editingId === loc.id ? (
              <>
                <input
                  value={editForm.name}
                  onChange={e => setEditForm({ ...editForm, name: e.target.value })}
                  className="flex-1 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-lg px-3 py-1.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                  placeholder="Nombre"
                />
                <input
                  type="number"
                  value={editForm.price}
                  onChange={e => setEditForm({ ...editForm, price: e.target.value })}
                  className="w-28 bg-[#1a1a1a] border border-[#ffd025]/30 rounded-lg px-3 py-1.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
                  placeholder="Precio"
                />
                <button onClick={() => handleUpdate(loc.id)} className="text-[#ffd025] hover:text-[#e5b81a] font-bold text-xs px-3 py-1.5 bg-[#ffd025]/10 rounded-lg">Guardar</button>
                <button onClick={() => setEditingId(null)} className="text-gray-500 hover:text-gray-300 text-xs px-2 py-1.5">✕</button>
              </>
            ) : (
              <>
                <span className={`w-2 h-2 rounded-full flex-shrink-0 ${loc.active ? "bg-green-400" : "bg-gray-600"}`} />
                <span className="flex-1 text-sm font-semibold text-white">{loc.name}</span>
                <span className="text-[#ffd025] font-black text-sm">${loc.price.toLocaleString("es-CL")}</span>
                <button
                  onClick={() => toggleActive(loc)}
                  className={`text-xs px-2.5 py-1 rounded-lg font-bold transition-all ${loc.active ? "bg-green-900/40 text-green-400 hover:bg-green-900/60" : "bg-gray-800 text-gray-500 hover:text-gray-300"}`}
                >
                  {loc.active ? "Activo" : "Inactivo"}
                </button>
                <button onClick={() => startEdit(loc)} className="text-gray-400 hover:text-[#ffd025] transition-colors">
                  <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/></svg>
                </button>
                <button onClick={() => handleDelete(loc.id)} className="text-gray-600 hover:text-red-400 transition-colors">
                  <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                </button>
              </>
            )}
          </div>
        ))}
      </div>
      <div className="border-t border-gray-800 pt-5 space-y-3">
        <p className="text-xs font-bold text-gray-400 uppercase">Nueva zona</p>
        <div className="flex gap-2">
          <input
            value={form.name}
            onChange={e => setForm({ ...form, name: e.target.value })}
            placeholder="Ej: Centro, La Florida…"
            className="flex-1 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-4 py-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
          />
          <input
            type="number"
            value={form.price}
            onChange={e => setForm({ ...form, price: e.target.value })}
            placeholder="Precio $"
            className="w-32 bg-[#141414] border border-[#ffd025]/20 rounded-xl px-4 py-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]"
          />
          <button
            onClick={handleCreate}
            disabled={!form.name.trim() || !form.price}
            className="px-4 py-2.5 bg-[#ffd025] text-[#141414] rounded-xl font-black text-sm hover:bg-[#e5b81a] transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            + Agregar
          </button>
        </div>
      </div>
    </div>
  );
}

function CommunityAdminPanel() {
  const { data: menu } = useGetMenu();
  const { data: menuSettings } = useGetMenu();
  const allProducts = (menu?.products ?? []) as Product[];
  const whatsapp = menuSettings?.settings?.whatsapp ?? "";

  const [socialSearch, setSocialSearch] = useState("");

  const normalizeStr = (s: string) =>
    s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
  const searchNorm = normalizeStr(socialSearch);

  const publishedProducts = allProducts.filter((p) => p.publishedSocial);
  const unpublishedProducts = allProducts.filter((p) => !p.publishedSocial);

  const filteredPublished = publishedProducts.filter(
    (p) => !searchNorm || normalizeStr(p.name).includes(searchNorm) || normalizeStr(p.category).includes(searchNorm),
  );
  const filteredUnpublished = unpublishedProducts.filter(
    (p) => !searchNorm || normalizeStr(p.name).includes(searchNorm) || normalizeStr(p.category).includes(searchNorm),
  );

  const [copiedId, setCopiedId] = useState<number | null>(null);
  const copyProductText = (p: Product) => {
    const lines = [
      `🛒 *${p.name}*`,
      `💰 $${Number(p.price).toLocaleString("es-CL")}`,
      p.category ? `📦 ${p.category}` : "",
      whatsapp ? `\n📲 Pedidos: wa.me/${whatsapp.replace(/\D/g, "")}` : "",
    ].filter(Boolean);
    navigator.clipboard.writeText(lines.join("\n"));
    setCopiedId(p.id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  return (
    <div className="space-y-6 animate-fade-in">
      <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div>
            <h2 className="text-xl font-black uppercase flex items-center gap-2">
              <Share2 size={20} className="text-[#ffd025]" /> Community Manager
            </h2>
            <p className="text-xs text-gray-400 mt-1">
              {publishedProducts.length} producto{publishedProducts.length !== 1 ? "s" : ""} publicado{publishedProducts.length !== 1 ? "s" : ""} · {unpublishedProducts.length} pendiente{unpublishedProducts.length !== 1 ? "s" : ""}
            </p>
          </div>
          <input
            type="text"
            value={socialSearch}
            onChange={(e) => setSocialSearch(e.target.value)}
            placeholder="Buscar producto..."
            className="w-full sm:w-64 bg-[#141414] border border-[#ffd025]/30 rounded-xl px-4 py-2.5 text-sm text-white placeholder-gray-500 focus:outline-none focus:border-[#ffd025]"
          />
        </div>

        {filteredPublished.length === 0 && filteredUnpublished.length === 0 && (
          <p className="text-center text-gray-500 py-10">No se encontraron productos.</p>
        )}

        {filteredPublished.length > 0 && (
          <div className="mb-8">
            <h3 className="text-xs font-black uppercase text-green-400 tracking-widest mb-4 flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-green-400 inline-block" />
              Publicados en redes ({filteredPublished.length})
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {filteredPublished.map((p) => (
                <div key={p.id} className="bg-[#141414] rounded-2xl border border-green-500/20 flex gap-3 p-3 items-start">
                  {p.image ? (
                    <img
                      src={p.image}
                      alt={p.name}
                      className="w-16 h-16 rounded-xl object-cover flex-shrink-0 bg-[#252525]"
                      onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
                    />
                  ) : (
                    <div className="w-16 h-16 rounded-xl bg-[#252525] flex-shrink-0 flex items-center justify-center">
                      <ImageIcon size={20} className="text-gray-600" />
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    <p className="font-bold text-sm text-white truncate">{p.name}</p>
                    <p className="text-xs text-gray-400 truncate">{p.category}{p.aisle ? ` · ${p.aisle}` : ""}</p>
                    <p className="text-sm font-black text-[#ffd025] mt-1">${Number(p.price).toLocaleString("es-CL")}</p>
                  </div>
                  <button
                    onClick={() => copyProductText(p)}
                    title="Copiar texto para redes"
                    className={`p-2 rounded-xl transition-colors flex-shrink-0 ${copiedId === p.id ? "text-green-400 bg-green-500/15" : "text-gray-500 hover:text-[#ffd025] hover:bg-[#ffd025]/10"}`}
                  >
                    {copiedId === p.id ? <CheckCircle size={15} /> : <Copy size={15} />}
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        {filteredUnpublished.length > 0 && (
          <div>
            <h3 className="text-xs font-black uppercase text-gray-500 tracking-widest mb-4 flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-gray-600 inline-block" />
              Pendientes de publicar ({filteredUnpublished.length})
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {filteredUnpublished.map((p) => (
                <div key={p.id} className="bg-[#141414] rounded-2xl border border-gray-800 flex gap-3 p-3 items-start opacity-50">
                  {p.image ? (
                    <img
                      src={p.image}
                      alt={p.name}
                      className="w-16 h-16 rounded-xl object-cover flex-shrink-0 bg-[#252525] grayscale"
                      onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
                    />
                  ) : (
                    <div className="w-16 h-16 rounded-xl bg-[#252525] flex-shrink-0 flex items-center justify-center">
                      <ImageIcon size={20} className="text-gray-600" />
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    <p className="font-bold text-sm text-white truncate">{p.name}</p>
                    <p className="text-xs text-gray-400 truncate">{p.category}{p.aisle ? ` · ${p.aisle}` : ""}</p>
                    <p className="text-sm font-black text-[#ffd025] mt-1">${Number(p.price).toLocaleString("es-CL")}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function OrdersAdminPanel({ role }: { role: "full" | "delivery" }) {
  const queryClient = useQueryClient();
  const { data: menu } = useGetMenu();
  const storeName = menu?.settings?.pageTitle || "Urban Bite";
  const allProducts = (menu?.products ?? []) as Product[];
  const { data: orders = [], isLoading, refetch: refetchOrders } = useListOrders();

  const patchOrderCache = (id: number, patch: Partial<Order>) =>
    queryClient.setQueryData(getListOrdersQueryKey(), (old: any) =>
      Array.isArray(old) ? old.map((o: Order) => (o.id === id ? { ...o, ...patch } : o)) : old,
    );

  useEffect(() => {
    const interval = setInterval(() => {
      refetchOrders();
    }, 15_000);
    return () => clearInterval(interval);
  }, [refetchOrders]);
  const deleteOrderMut = useDeleteOrder();
  const updateStatusMut = useUpdateOrderStatus();

  type EditItem = { name: string; quantity: number; price: number; selectedOption?: string };
  const [editingOrder, setEditingOrder] = useState<Order | null>(null);
  const [editItems, setEditItems] = useState<EditItem[]>([]);
  const [editSearch, setEditSearch] = useState("");
  const [editNote, setEditNote] = useState("");
  const [panelToast, setPanelToast] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setPanelToast(msg);
    setTimeout(() => setPanelToast(null), 3000);
  };

  const openEdit = (order: Order) => {
    const items = (order.items ?? []) as EditItem[];
    setEditItems(items.map((it) => ({ ...it })));
    setEditNote((order as Order & { modificationNote?: string | null }).modificationNote ?? "");
    setEditSearch("");
    setEditingOrder(order);
  };

  const handleSaveEdit = () => {
    if (!editingOrder) return;
    const newTotal = editItems.reduce((s, it) => s + it.price * it.quantity, 0);
    patchOrderCache(editingOrder.id, { items: editItems as any, total: String(newTotal), modificationNote: editNote || null } as any);
    setEditingOrder(null);
    updateStatusMut.mutate(
      {
        id: editingOrder.id,
        data: {
          status: editingOrder.status,
          items: editItems,
          modificationNote: editNote || undefined,
          total: newTotal,
        } as Parameters<typeof updateStatusMut.mutate>[0]["data"],
      },
      {
        onError: () => {
          refetchOrders();
          showToast("Error al guardar los cambios. Intenta de nuevo.");
        },
      },
    );
  };

  const handleSetCaja = (order: Order, caja: string | null) => {
    patchOrderCache(order.id, { caja: caja ?? undefined } as any);
    updateStatusMut.mutate({
      id: order.id,
      data: { status: order.status, caja: caja ?? undefined } as Parameters<typeof updateStatusMut.mutate>[0]["data"],
    }, { onError: () => refetchOrders() });
  };

  const knownIdsRef = useRef<Set<number> | null>(null);
  const notifPermissionRef = useRef<NotificationPermission>("default");

  useEffect(() => {
    if ("Notification" in window && Notification.permission === "default") {
      Notification.requestPermission().then((perm) => {
        notifPermissionRef.current = perm;
      });
    } else if ("Notification" in window) {
      notifPermissionRef.current = Notification.permission;
    }
  }, []);

  useEffect(() => {
    if (isLoading || orders.length === 0) return;
    const currentIds = new Set(orders.map((o: Order) => o.id));

    if (knownIdsRef.current === null) {
      knownIdsRef.current = currentIds;
      return;
    }

    const newOrders = orders.filter(
      (o: Order) => !knownIdsRef.current!.has(o.id),
    );
    knownIdsRef.current = currentIds;

    if (newOrders.length === 0) return;

    if (
      "Notification" in window &&
      (notifPermissionRef.current === "granted" ||
        Notification.permission === "granted")
    ) {
      newOrders.forEach((o: Order) => {
        new Notification(`🛒 Nuevo pedido — ${storeName}`, {
          body: `${o.customerName} · $${Number(o.total).toLocaleString("es-CL")}`,
          icon: "/favicon.ico",
        });
      });
    }
  }, [orders, isLoading, storeName]);

  const handleDelete = (id: number) => {
    if (!confirm("¿Eliminar este pedido del registro?")) return;
    queryClient.setQueryData(getListOrdersQueryKey(), (old: any) =>
      Array.isArray(old) ? old.filter((o: Order) => o.id !== id) : old,
    );
    deleteOrderMut.mutate({ id }, { onError: () => refetchOrders() });
  };

  const handleAnular = (id: number) => {
    try {
      if (typeof window !== "undefined" && window.confirm && !window.confirm("¿Anular este pedido? Quedará registrado pero no contará en las ventas.")) return;
    } catch {}
    patchOrderCache(id, { status: "Anulado" });
    updateStatusMut.mutate(
      { id, data: { status: "Anulado" } },
      { onError: () => { refetchOrders(); showToast("Error al anular el pedido. Intenta de nuevo."); } },
    );
  };

  const buildStatusMessage = (customerName: string, status: string) => {
    const first = customerName.split(" ")[0] || customerName;
    switch (status) {
      case "Confirmado y en Preparación":
        return `¡Hola ${first}! 👋 Tu pedido en *${storeName}* fue *confirmado* y ya lo estamos alistando 🍾\n\nEn breve te avisamos cuando esté listo.`;
      case "Listo para Retirar":
        return `¡Hola ${first}! 🍻 Tu pedido en *${storeName}* ya está *listo para retirar* 🏪✅\n\nTe esperamos en el local con todo listo. ¡Salud! 🥃`;
      case "Delivery en Camino":
        return `¡Hola ${first}! 🛵 Tu pedido de *${storeName}* va *en camino* con nuestro delivery 🍾\n\nEn unos minutos llega a tu dirección. ¡Gracias por tu preferencia! 🥂🎊`;
      default:
        return "";
    }
  };

  const handleSetStatus = (
    id: number,
    status: string,
    customerName: string,
    phone: string,
    sendsWhatsapp: boolean,
  ) => {
    if (sendsWhatsapp) {
      const cleanedPhone = (phone ?? "").replace(/\D/g, "");
      if (!cleanedPhone) {
        showToast("Este pedido no tiene número de teléfono. No se puede enviar el mensaje.");
        return;
      }
      const message = buildStatusMessage(customerName, status);
      try {
        window.open(
          `https://api.whatsapp.com/send?phone=${cleanedPhone}&text=${encodeURIComponent(message)}`,
          "_blank",
        );
      } catch {}
    }
    patchOrderCache(id, { status });
    updateStatusMut.mutate(
      { id, data: { status } },
      { onError: () => { refetchOrders(); showToast("Error al actualizar el estado. Intenta de nuevo."); } },
    );
  };

  const normalizeStr = (s: string) =>
    s.toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");

  return (
    <>
    {editingOrder && (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
        <div className="bg-[#1a1a1a] rounded-3xl border border-[#ffd025]/30 w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl">
          <div className="flex items-center justify-between p-5 border-b border-gray-800">
            <h3 className="text-base font-black uppercase flex items-center gap-2">
              <Pencil size={16} className="text-[#ffd025]" />
              Editar Pedido #{editingOrder.id} — {editingOrder.customerName}
            </h3>
            <button
              type="button"
              onClick={() => setEditingOrder(null)}
              className="p-2 text-gray-500 hover:text-white hover:bg-gray-800 rounded-xl transition-colors"
            >
              <X size={16} />
            </button>
          </div>
          <div className="flex-1 overflow-y-auto p-5 space-y-5">
            <div>
              <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                Productos del pedido
              </p>
              <ul className="space-y-1.5">
                {editItems.map((it, idx) => (
                  <li
                    key={idx}
                    className="flex items-center gap-3 bg-[#141414] border border-gray-800 rounded-xl px-3 py-2"
                  >
                    <div className="flex items-center gap-1.5 mr-1">
                      <button
                        type="button"
                        onClick={() =>
                          setEditItems(
                            editItems.map((x, i) =>
                              i === idx ? { ...x, quantity: Math.max(1, x.quantity - 1) } : x,
                            ),
                          )
                        }
                        className="w-5 h-5 rounded-full bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold flex items-center justify-center transition-colors"
                      >
                        −
                      </button>
                      <span className="text-[#ffd025] font-black text-sm w-5 text-center">
                        {it.quantity}
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          setEditItems(
                            editItems.map((x, i) =>
                              i === idx ? { ...x, quantity: x.quantity + 1 } : x,
                            ),
                          )
                        }
                        className="w-5 h-5 rounded-full bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold flex items-center justify-center transition-colors"
                      >
                        +
                      </button>
                    </div>
                    <span className="flex-1 text-sm text-white">
                      {it.name}
                      {it.selectedOption && (
                        <span className="text-gray-500"> ({it.selectedOption})</span>
                      )}
                    </span>
                    <span className="text-gray-400 text-sm">
                      ${(it.price * it.quantity).toLocaleString("es-CL")}
                    </span>
                    <button
                      type="button"
                      onClick={() => setEditItems(editItems.filter((_, i) => i !== idx))}
                      className="p-1 text-red-500 hover:text-red-400 hover:bg-red-500/10 rounded-lg transition-colors"
                      title="Quitar producto"
                    >
                      <X size={13} />
                    </button>
                  </li>
                ))}
              </ul>
              {editItems.length === 0 && (
                <p className="text-gray-600 text-sm text-center py-4">
                  Sin productos — agrega uno desde el buscador
                </p>
              )}
              {editItems.length > 0 && (
                <div className="mt-2 text-right text-sm">
                  <span className="text-gray-500">Nuevo total: </span>
                  <span className="text-[#ffd025] font-black">
                    ${editItems.reduce((s, it) => s + it.price * it.quantity, 0).toLocaleString("es-CL")}
                  </span>
                </div>
              )}
            </div>

            <div>
              <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                Agregar / reemplazar producto
              </p>
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none">
                  🔍
                </span>
                <input
                  type="text"
                  placeholder="Buscar en el catálogo..."
                  value={editSearch}
                  onChange={(e) => setEditSearch(e.target.value)}
                  className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl pl-9 py-2 text-white text-sm focus:border-[#ffd025] focus:outline-none placeholder:text-gray-600"
                />
              </div>
              {editSearch.trim() && (() => {
                const q = normalizeStr(editSearch.trim());
                const words = q.split(/\s+/).filter(Boolean);
                const results = allProducts
                  .filter((p) => {
                    const haystack = normalizeStr(
                      [p.name, p.category, p.aisle, p.subcategory ?? ""].join(" "),
                    );
                    return words.every((w) => haystack.includes(w));
                  })
                  .slice(0, 12);
                return (
                  <div className="mt-1.5 max-h-44 overflow-y-auto space-y-1 rounded-xl border border-gray-800 bg-[#141414] p-1">
                    {results.length === 0 ? (
                      <p className="text-gray-600 text-sm text-center py-3">Sin resultados</p>
                    ) : (
                      results.map((p) => (
                        <button
                          key={p.id}
                          type="button"
                          onClick={() => {
                            setEditItems([
                              ...editItems,
                              { name: p.name, quantity: 1, price: p.price },
                            ]);
                            setEditSearch("");
                          }}
                          className="w-full text-left flex items-center justify-between gap-2 hover:bg-[#222] rounded-lg px-3 py-1.5 transition-colors"
                        >
                          <span className="text-sm text-white">{p.name}</span>
                          <span className="text-[#ffd025] font-bold text-sm whitespace-nowrap">
                            ${p.price.toLocaleString("es-CL")}
                          </span>
                        </button>
                      ))
                    )}
                  </div>
                );
              })()}
            </div>

            <div>
              <label className="text-[10px] uppercase font-bold text-gray-500 mb-2 block tracking-widest">
                Motivo del cambio (opcional)
              </label>
              <textarea
                rows={2}
                placeholder="Ej: No había Cerveza Austral, se reemplazó por Cristal 500ml..."
                value={editNote}
                onChange={(e) => setEditNote(e.target.value)}
                className="w-full bg-[#141414] border border-[#ffd025]/30 rounded-xl p-3 text-white text-sm focus:border-[#ffd025] focus:outline-none resize-none placeholder:text-gray-600"
              />
            </div>
          </div>
          <div className="flex gap-3 p-5 border-t border-gray-800">
            <button
              type="button"
              onClick={() => setEditingOrder(null)}
              className="flex-1 py-2.5 rounded-xl border border-gray-700 text-gray-400 hover:text-white hover:border-gray-500 text-sm font-bold transition-colors"
            >
              Cancelar
            </button>
            <button
              type="button"
              onClick={handleSaveEdit}
              disabled={updateStatusMut.isPending || editItems.length === 0}
              className="flex-1 py-2.5 bg-[#ffd025] text-[#141414] rounded-xl font-bold text-sm uppercase hover:bg-[#e5b81a] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Guardar cambios
            </button>
          </div>
        </div>
      </div>
    )}
    <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20 animate-fade-in">
      <h2 className="text-xl font-black uppercase mb-6 flex items-center gap-2">
        <ClipboardList size={20} className="text-[#ffd025]" /> Registro de Pedidos
        <span className="ml-2 text-xs font-bold text-[#ffd025] bg-[#ffd025]/10 px-2.5 py-1 rounded-full">
          {orders.length}
        </span>
      </h2>

      {isLoading ? (
        <p className="text-gray-500 text-sm">Cargando pedidos...</p>
      ) : orders.length === 0 ? (
        <div className="text-center py-12 opacity-60">
          <ClipboardList size={40} className="mx-auto mb-3 text-[#ffd025]/40" />
          <p className="text-sm uppercase font-bold text-gray-400">
            Aún no hay pedidos registrados
          </p>
        </div>
      ) : (
        <div className="space-y-4">
          {orders.map((order: Order) => {
            const date = new Date(order.createdAt);
            const items = (order.items ?? []) as Array<{
              name: string;
              quantity: number;
              price: number;
              selectedOption?: string;
            }>;
            return (
              <div
                key={order.id}
                className={`border rounded-2xl p-5 transition-colors ${
                  order.status === "Anulado"
                    ? "bg-[#1a1010] border-red-900/50 opacity-60"
                    : order.status === "Entregado"
                    ? "bg-[#0d1a0f] border-green-700/40 opacity-70"
                    : "bg-[#141414] border-gray-800 hover:border-[#ffd025]/40"
                }`}
              >
                <div className="flex items-start justify-between gap-3 mb-3">
                  <div>
                    <div className="flex items-center gap-2 text-white font-bold text-base">
                      <User size={16} className="text-[#ffd025]" />
                      {order.customerName}
                    </div>
                    <div className="flex items-start gap-2 mt-1.5 text-sm text-gray-300">
                      <MapPin size={14} className="text-[#ffd025] mt-0.5 flex-shrink-0" />
                      <span>{order.address}</span>
                    </div>
                    {order.phone && (
                      <div className="flex items-center gap-2 mt-1.5 text-sm text-gray-400">
                        <Phone size={13} className="text-[#ffd025]" />
                        <span>{order.phone}</span>
                      </div>
                    )}
                    <p className="text-[10px] uppercase tracking-widest text-gray-600 mt-2">
                      {date.toLocaleString("es-CL", {
                        day: "2-digit",
                        month: "2-digit",
                        year: "numeric",
                        hour: "2-digit",
                        minute: "2-digit",
                      })}
                    </p>
                  </div>
                  <div className="flex flex-col items-end gap-2">
                    <span className="text-xl font-black text-[#ffd025]">
                      ${order.total.toLocaleString("es-CL")}
                    </span>
                    {order.discountCode && order.discountAmount && order.subtotal && (
                      <div className="flex flex-col items-end gap-0.5">
                        <span className="text-[10px] font-bold uppercase tracking-wide text-green-400 bg-green-500/10 border border-green-500/25 px-2 py-0.5 rounded-full">
                          🏷️ {order.discountCode}
                        </span>
                        <span className="text-[11px] text-gray-400 line-through">
                          ${order.subtotal.toLocaleString("es-CL")}
                        </span>
                        <span className="text-[11px] text-green-400 font-semibold">
                          −${order.discountAmount.toLocaleString("es-CL")}
                        </span>
                      </div>
                    )}
                    {order.status === "Anulado" && (
                      <span className="text-[10px] font-black uppercase tracking-widest text-red-400 bg-red-600/15 border border-red-600/30 px-2 py-0.5 rounded-full flex items-center gap-1">
                        <Ban size={10} /> Anulado
                      </span>
                    )}
                    {order.status === "Entregado" && (
                      <span className="text-[10px] font-black uppercase tracking-widest text-green-400 bg-green-600/15 border border-green-600/30 px-2 py-0.5 rounded-full">
                        Venta cerrada
                      </span>
                    )}
                    {order.status !== "Anulado" && (
                      <button
                        onClick={() => openEdit(order)}
                        className="p-1.5 text-gray-600 hover:text-[#ffd025] hover:bg-[#ffd025]/10 rounded-lg transition-colors"
                        aria-label="Editar pedido"
                        title="Editar productos del pedido"
                      >
                        <Pencil size={15} />
                      </button>
                    )}
                    {order.status !== "Anulado" && (
                      <button
                        onClick={() => handleAnular(order.id)}
                        disabled={updateStatusMut.isPending}
                        className="p-1.5 text-gray-600 hover:text-orange-400 hover:bg-orange-500/10 rounded-lg transition-colors disabled:opacity-40"
                        aria-label="Anular pedido"
                        title="Anular pedido (queda registrado pero no cuenta en ventas)"
                      >
                        <Ban size={15} />
                      </button>
                    )}
                    {role === "full" && (
                      <button
                        onClick={() => handleDelete(order.id)}
                        className="p-1.5 text-gray-600 hover:text-red-500 hover:bg-red-500/10 rounded-lg transition-colors"
                        aria-label="Eliminar pedido"
                      >
                        <Trash2 size={15} />
                      </button>
                    )}
                  </div>
                </div>

                {order.status !== "Anulado" && (
                  <div className="mt-2 mb-1 flex items-center gap-2 flex-wrap">
                    <CreditCard size={12} className="text-gray-600" />
                    <span className="text-[10px] uppercase font-bold text-gray-600 tracking-widest">Cobrado desde:</span>
                    {[null, "Caja 1", "Caja 2"].map((c) => {
                      const orderCaja = (order as Order & { caja?: string | null }).caja ?? null;
                      const isActive = orderCaja === c;
                      return (
                        <button
                          key={c ?? "ninguna"}
                          type="button"
                          onClick={() => handleSetCaja(order, c)}
                          disabled={updateStatusMut.isPending}
                          className={`text-[10px] px-2.5 py-0.5 rounded-full font-bold border transition-colors disabled:opacity-50 ${
                            isActive
                              ? c === null
                                ? "bg-gray-700 border-gray-500 text-gray-300"
                                : c === "Caja 1"
                                ? "bg-blue-600 border-blue-400 text-white"
                                : "bg-purple-600 border-purple-400 text-white"
                              : "bg-transparent border-gray-800 text-gray-600 hover:border-gray-600 hover:text-gray-400"
                          }`}
                        >
                          {c ?? "Sin asignar"}
                        </button>
                      );
                    })}
                  </div>
                )}

                <div className="border-t border-gray-800 pt-3 mt-3">
                  <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                    Pedido
                  </p>
                  <ul className="space-y-1">
                    {items.map((it, idx) => (
                      <li key={idx} className="flex justify-between text-sm text-gray-300">
                        <span>
                          <span className="text-[#ffd025] font-bold mr-2">
                            {it.quantity}x
                          </span>
                          {it.name}
                          {it.selectedOption && (
                            <span className="text-gray-500"> ({it.selectedOption})</span>
                          )}
                        </span>
                        <span className="text-gray-400">
                          ${(it.price * it.quantity).toLocaleString("es-CL")}
                        </span>
                      </li>
                    ))}
                  </ul>
                </div>

                {order.notes && (
                  <div className="mt-3 pt-3 border-t border-gray-800 text-xs text-gray-400">
                    <span className="font-bold text-gray-500 uppercase mr-2">Notas:</span>
                    {order.notes}
                  </div>
                )}

                {(order as Order & { modificationNote?: string | null }).modificationNote && (
                  <div className="mt-2 pt-2 border-t border-amber-900/30 text-xs text-amber-400/80 bg-amber-900/10 rounded-xl px-3 py-2">
                    <span className="font-bold text-amber-500 uppercase mr-2">✏️ Cambio:</span>
                    {(order as Order & { modificationNote?: string | null }).modificationNote}
                  </div>
                )}

                {order.status !== "Anulado" && (
                  <div className="mt-4 pt-3 border-t border-gray-800">
                    <p className="text-[10px] uppercase font-bold text-gray-500 mb-2 tracking-widest">
                      Estado del pedido
                    </p>
                    <div className="flex flex-wrap gap-2">
                      {ORDER_STATUSES.map((s) => {
                        const isActive = order.status === s.value;
                        return (
                          <button
                            key={s.value}
                            type="button"
                            onClick={() =>
                              handleSetStatus(
                                order.id,
                                s.value,
                                order.customerName,
                                order.phone ?? "",
                                s.whatsapp,
                              )
                            }
                            disabled={updateStatusMut.isPending}
                            className={`text-[11px] md:text-xs font-bold uppercase tracking-wide px-3 py-1.5 rounded-full border transition-all disabled:opacity-50 ${
                              isActive ? s.activeColor : s.color + " hover:brightness-125"
                            } ${s.value === "Entregado" && !isActive ? "ring-1 ring-green-600/30" : ""}`}
                            data-testid={`button-order-status-${order.id}-${s.value}`}
                          >
                            {s.label}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
                {order.status === "Anulado" && (
                  <div className="mt-4 pt-3 border-t border-red-900/30">
                    <p className="text-[10px] uppercase font-bold text-red-900 tracking-widest">
                      Pedido anulado · no contabilizado en ventas
                    </p>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
    </>
  );
}

const DOW_NAMES = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

function StatsAdminPanel() {
  const [period, setPeriod] = useState<"weekly" | "monthly">("weekly");
  const [isGenerating, setIsGenerating] = useState(false);
  const visitsQ = useGetVisitStats();
  const salesQ = useGetSalesStats({ period });
  const { data: menu } = useGetMenu();
  const storeName = menu?.settings?.pageTitle || "Urban Bite";

  const visits = visitsQ.data;
  const sales = salesQ.data;

  const handleDownloadPDF = async () => {
    setIsGenerating(true);
    try {
      const [freshVisits, freshSales] = await Promise.all([
        visitsQ.refetch(),
        salesQ.refetch(),
      ]);
      generateStatsPDF(storeName, period, freshVisits.data, freshSales.data);
    } finally {
      setIsGenerating(false);
    }
  };

  const maxVisitHour = visits
    ? Math.max(1, ...visits.hourlyDistribution.map((b) => b.count))
    : 1;
  const maxSalesHour = sales
    ? Math.max(1, ...sales.hourlyDistribution.map((b) => b.orders))
    : 1;
  const maxDow = sales
    ? Math.max(1, ...sales.dayOfWeekDistribution.map((b) => b.orders))
    : 1;
  const maxTopQty = sales?.topProducts.length
    ? Math.max(1, ...sales.topProducts.map((p) => p.qty))
    : 1;

  const peakVisitHour = visits?.hourlyDistribution.reduce(
    (best, cur) => (cur.count > best.count ? cur : best),
    { hour: 0, count: 0 },
  );
  const peakSalesHour = sales?.hourlyDistribution.reduce(
    (best, cur) => (cur.orders > best.orders ? cur : best),
    { hour: 0, orders: 0, revenue: 0 },
  );
  const peakDow = sales?.dayOfWeekDistribution.reduce(
    (best, cur) => (cur.orders > best.orders ? cur : best),
    { dayOfWeek: 0, orders: 0, revenue: 0 },
  );

  return (
    <div className="space-y-6 animate-fade-in">

      <div className="bg-[#1a1a1a] p-5 rounded-3xl border border-[#ffd025]/20 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-black uppercase flex items-center gap-2">
            <BarChart3 size={20} className="text-[#ffd025]" /> Infografía PDF
          </h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Descarga un resumen visual de visitas y ventas en formato PDF.
          </p>
        </div>
        <button
          onClick={handleDownloadPDF}
          disabled={isGenerating || (!visits && !sales)}
          className="flex items-center gap-2 px-5 py-3 bg-gradient-to-r from-[#ffd025] to-orange-400 text-[#141414] rounded-2xl font-black uppercase text-sm hover:brightness-110 transition-all shadow-lg shadow-[#ffd025]/20 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          <Download size={18} />
          {isGenerating ? "Generando..." : "Descargar Infografía"}
        </button>
      </div>

      <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
        <h2 className="text-xl font-black uppercase mb-1 flex items-center gap-2">
          <Eye size={20} className="text-[#ffd025]" /> Visitas a la página
        </h2>
        <p className="text-xs text-gray-500 mb-5">
          Datos basados en la IP de los visitantes (ciudad y horario).
        </p>

        {visitsQ.isLoading || !visits ? (
          <p className="text-gray-500 text-sm">Cargando estadísticas de visitas...</p>
        ) : (
          <>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
              <StatCard label="Total" value={visits.total} icon={<Eye size={16} />} />
              <StatCard label="Hoy" value={visits.today} icon={<Clock size={16} />} />
              <StatCard label="7 días" value={visits.last7Days} icon={<Calendar size={16} />} />
              <StatCard label="30 días" value={visits.last30Days} icon={<TrendingUp size={16} />} />
            </div>

            {peakVisitHour && peakVisitHour.count > 0 && (
              <div className="mb-5 text-xs uppercase tracking-wider text-gray-400">
                Hora pico (últimos 7 días):{" "}
                <span className="text-[#ffd025] font-bold">
                  {String(peakVisitHour.hour).padStart(2, "0")}:00 ·{" "}
                  {peakVisitHour.count} visitas
                </span>
              </div>
            )}

            <div className="mb-6">
              <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                <Clock size={14} className="text-[#ffd025]" /> Visitas por hora (últimos 7 días)
              </h3>
              <div className="flex items-end gap-1 h-32 bg-[#141414] rounded-2xl p-3 border border-gray-800">
                {visits.hourlyDistribution.map((b) => (
                  <div key={b.hour} className="flex-1 flex flex-col items-center gap-1 group">
                    <div className="text-[9px] text-gray-600 group-hover:text-[#ffd025]">
                      {b.count > 0 ? b.count : ""}
                    </div>
                    <div
                      className="w-full bg-gradient-to-t from-[#ffd025]/30 to-[#ffd025] rounded-t transition-all"
                      style={{ height: `${(b.count / maxVisitHour) * 90}%`, minHeight: b.count > 0 ? 4 : 0 }}
                      title={`${String(b.hour).padStart(2, "0")}:00 - ${b.count} visitas`}
                    />
                    <div className="text-[9px] text-gray-500">{b.hour}</div>
                  </div>
                ))}
              </div>
            </div>

            <div className="grid md:grid-cols-2 gap-4">
              <div>
                <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                  <MapPin size={14} className="text-[#ffd025]" /> Top ciudades (30 días)
                </h3>
                {visits.topCities.length === 0 ? (
                  <p className="text-xs text-gray-600">Sin datos suficientes todavía.</p>
                ) : (
                  <ul className="space-y-1.5">
                    {visits.topCities.map((c, i) => (
                      <li
                        key={`${c.city}-${i}`}
                        className="flex justify-between items-center text-sm bg-[#141414] px-3 py-2 rounded-xl border border-gray-800"
                      >
                        <span className="text-gray-200">
                          <span className="text-[#ffd025] font-bold mr-2">{i + 1}.</span>
                          {c.city}
                          {c.country && (
                            <span className="text-gray-500 text-xs ml-1">· {c.country}</span>
                          )}
                        </span>
                        <span className="font-bold text-[#ffd025]">{c.count}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
              <div>
                <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                  <Globe size={14} className="text-[#ffd025]" /> Top países (30 días)
                </h3>
                {visits.topCountries.length === 0 ? (
                  <p className="text-xs text-gray-600">Sin datos suficientes todavía.</p>
                ) : (
                  <ul className="space-y-1.5">
                    {visits.topCountries.map((c, i) => (
                      <li
                        key={`${c.country}-${i}`}
                        className="flex justify-between items-center text-sm bg-[#141414] px-3 py-2 rounded-xl border border-gray-800"
                      >
                        <span className="text-gray-200">
                          <span className="text-[#ffd025] font-bold mr-2">{i + 1}.</span>
                          {c.country}
                        </span>
                        <span className="font-bold text-[#ffd025]">{c.count}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>
          </>
        )}
      </div>

      <div className="bg-[#1a1a1a] p-6 rounded-3xl border border-[#ffd025]/20">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-1">
          <h2 className="text-xl font-black uppercase flex items-center gap-2">
            <BarChart3 size={20} className="text-[#ffd025]" /> Reporte de Ventas
          </h2>
          <div className="inline-flex bg-[#141414] rounded-full p-1 border border-gray-800">
            <button
              onClick={() => setPeriod("weekly")}
              className={`px-4 py-1.5 text-xs font-bold uppercase rounded-full transition-colors ${
                period === "weekly"
                  ? "bg-[#ffd025] text-[#141414]"
                  : "text-gray-400 hover:text-white"
              }`}
            >
              Semanal
            </button>
            <button
              onClick={() => setPeriod("monthly")}
              className={`px-4 py-1.5 text-xs font-bold uppercase rounded-full transition-colors ${
                period === "monthly"
                  ? "bg-[#ffd025] text-[#141414]"
                  : "text-gray-400 hover:text-white"
              }`}
            >
              Mensual
            </button>
          </div>
        </div>
        <p className="text-xs text-gray-500 mb-5">
          {period === "weekly" ? "Últimos 7 días" : "Últimos 30 días"} · Lo más vendido y horarios pico.
        </p>

        {salesQ.isLoading || !sales ? (
          <p className="text-gray-500 text-sm">Cargando reporte de ventas...</p>
        ) : (
          <>
            <div className="grid grid-cols-2 gap-3 mb-6">
              <StatCard
                label="Pedidos entregados"
                value={sales.totalOrders}
                icon={<ClipboardList size={16} />}
              />
              <StatCard
                label="Ingresos"
                value={`$${sales.totalRevenue.toLocaleString("es-CL")}`}
                icon={<TrendingUp size={16} />}
              />
            </div>
            {sales.cancelledOrders > 0 && (
              <div className="flex items-center gap-2 mb-6 px-3 py-2.5 rounded-xl border border-red-900/40 bg-red-950/20 text-xs text-red-400">
                <Ban size={13} className="flex-shrink-0" />
                <span>
                  <span className="font-black">{sales.cancelledOrders}</span>{" "}
                  {sales.cancelledOrders === 1 ? "pedido anulado" : "pedidos anulados"} en este período · no incluidos en los ingresos
                </span>
              </div>
            )}

            <div className="grid md:grid-cols-2 gap-3 mb-6 text-xs uppercase tracking-wider text-gray-400">
              {peakSalesHour && peakSalesHour.orders > 0 && (
                <div className="bg-[#141414] px-3 py-2 rounded-xl border border-gray-800">
                  Hora pico:{" "}
                  <span className="text-[#ffd025] font-bold">
                    {String(peakSalesHour.hour).padStart(2, "0")}:00 ·{" "}
                    {peakSalesHour.orders} pedidos
                  </span>
                </div>
              )}
              {peakDow && peakDow.orders > 0 && (
                <div className="bg-[#141414] px-3 py-2 rounded-xl border border-gray-800">
                  Día pico:{" "}
                  <span className="text-[#ffd025] font-bold">
                    {DOW_NAMES[peakDow.dayOfWeek]} · {peakDow.orders} pedidos
                  </span>
                </div>
              )}
            </div>

            <div className="mb-6">
              <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                <Flame size={14} className="text-[#ffd025]" /> Lo más vendido
              </h3>
              {sales.topProducts.length === 0 ? (
                <p className="text-xs text-gray-600">
                  Aún no hay ventas en este período.
                </p>
              ) : (
                <ul className="space-y-2">
                  {sales.topProducts.map((p, i) => (
                    <li key={`${p.name}-${i}`} className="space-y-1">
                      <div className="flex justify-between items-center text-sm">
                        <span className="text-gray-200 truncate pr-2">
                          <span className="text-[#ffd025] font-bold mr-2">{i + 1}.</span>
                          {p.name}
                        </span>
                        <span className="font-bold text-[#ffd025] whitespace-nowrap">
                          {p.qty}u · ${p.revenue.toLocaleString("es-CL")}
                        </span>
                      </div>
                      <div className="h-1.5 bg-[#141414] rounded-full overflow-hidden">
                        <div
                          className="h-full bg-gradient-to-r from-[#ffd025] to-orange-500 rounded-full"
                          style={{ width: `${(p.qty / maxTopQty) * 100}%` }}
                        />
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </div>

            <div className="grid md:grid-cols-2 gap-6">
              <div>
                <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                  <Clock size={14} className="text-[#ffd025]" /> Pedidos por hora
                </h3>
                <div className="flex items-end gap-1 h-32 bg-[#141414] rounded-2xl p-3 border border-gray-800">
                  {sales.hourlyDistribution.map((b) => (
                    <div key={b.hour} className="flex-1 flex flex-col items-center gap-1 group">
                      <div className="text-[9px] text-gray-600 group-hover:text-[#ffd025]">
                        {b.orders > 0 ? b.orders : ""}
                      </div>
                      <div
                        className="w-full bg-gradient-to-t from-orange-500/40 to-[#ffd025] rounded-t"
                        style={{
                          height: `${(b.orders / maxSalesHour) * 90}%`,
                          minHeight: b.orders > 0 ? 4 : 0,
                        }}
                        title={`${String(b.hour).padStart(2, "0")}:00 - ${b.orders} pedidos`}
                      />
                      <div className="text-[9px] text-gray-500">{b.hour}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div>
                <h3 className="text-xs uppercase font-bold text-gray-400 mb-3 flex items-center gap-2">
                  <Calendar size={14} className="text-[#ffd025]" /> Pedidos por día de la semana
                </h3>
                <div className="flex items-end gap-2 h-32 bg-[#141414] rounded-2xl p-3 border border-gray-800">
                  {sales.dayOfWeekDistribution.map((b) => (
                    <div
                      key={b.dayOfWeek}
                      className="flex-1 flex flex-col items-center gap-1 group"
                    >
                      <div className="text-[9px] text-gray-600 group-hover:text-[#ffd025]">
                        {b.orders > 0 ? b.orders : ""}
                      </div>
                      <div
                        className="w-full bg-gradient-to-t from-orange-500/40 to-[#ffd025] rounded-t"
                        style={{
                          height: `${(b.orders / maxDow) * 90}%`,
                          minHeight: b.orders > 0 ? 4 : 0,
                        }}
                        title={`${DOW_NAMES[b.dayOfWeek]} - ${b.orders} pedidos`}
                      />
                      <div className="text-[10px] text-gray-500 font-bold">
                        {DOW_NAMES[b.dayOfWeek].slice(0, 3)}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function CustomersAdminPanel() {
  const [customersData, setCustomersData] = useState<{
    registered: Array<{ id: number; email: string; name: string; phone: string; createdAt: string; orderCount: number; totalSpent: number }>;
    topBuyers: Array<{ name: string; phone: string; orderCount: number; totalSpent: number }>;
  } | null>(null);
  const [discounts, setDiscounts] = useState<Array<{
    id: number; code: string; type: string; amount: number; customerId: number | null;
    customerEmail: string | null; customerName: string | null; active: boolean;
    minOrder: number; usesLeft: number | null; expiresAt: string | null;
  }>>([]);
  const [loading, setLoading] = useState(true);
  const [newDiscount, setNewDiscount] = useState<{
    code: string; type: "percentage" | "fixed"; amount: number; customerId: number | null;
    minOrder: number; usesLeft: number | null; expiresAt: string;
  }>({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" });
  const [creating, setCreating] = useState(false);
  const [showCreateForm, setShowCreateForm] = useState(false);

  const loadData = async () => {
    setLoading(true);
    try {
      const [cRes, dRes] = await Promise.all([
        fetch("/api/admin/customers"),
        fetch("/api/admin/discounts"),
      ]);
      if (cRes.ok) setCustomersData(await cRes.json());
      if (dRes.ok) setDiscounts(await dRes.json());
    } catch {}
    setLoading(false);
  };

  useEffect(() => { loadData(); }, []);

  const createDiscount = async (e: FormEvent) => {
    e.preventDefault();
    if (!newDiscount.code.trim() || newDiscount.amount <= 0) return;
    setCreating(true);
    try {
      const r = await fetch("/api/admin/discounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...newDiscount, expiresAt: newDiscount.expiresAt || null }),
      });
      if (r.ok) {
        setNewDiscount({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" });
        setShowCreateForm(false);
        await loadData();
      }
    } catch {}
    setCreating(false);
  };

  const toggleDiscount = async (id: number, active: boolean) => {
    try {
      await fetch(`/api/admin/discounts/${id}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ active }) });
      setDiscounts(prev => prev.map(d => d.id === id ? { ...d, active } : d));
    } catch {}
  };

  const deleteDiscount = async (id: number) => {
    if (!confirm("¿Eliminar este descuento?")) return;
    try {
      await fetch(`/api/admin/discounts/${id}`, { method: "DELETE" });
      setDiscounts(prev => prev.filter(d => d.id !== id));
    } catch {}
  };

  if (loading) return (
    <div className="flex items-center justify-center py-20">
      <Loader2 size={32} className="animate-spin text-[#ffd025]" />
    </div>
  );

  return (
    <div className="space-y-8 animate-fade-in">
      {/* Top buyers */}
      <div className="bg-[#1a1a1a] rounded-3xl border border-[#ffd025]/20 p-6">
        <h2 className="text-xl font-black uppercase mb-4 flex items-center gap-2">
          <TrendingUp size={20} className="text-[#ffd025]" /> Top compradores
        </h2>
        {!customersData?.topBuyers.length ? (
          <p className="text-gray-500 text-sm">Aún no hay pedidos registrados.</p>
        ) : (
          <div className="space-y-2">
            {customersData.topBuyers.map((b, i) => (
              <div key={i} className="flex items-center justify-between bg-[#141414] rounded-2xl px-4 py-3">
                <div className="flex items-center gap-3">
                  <span className="text-[#ffd025] font-black text-lg w-7 shrink-0">#{i + 1}</span>
                  <div>
                    <p className="font-bold text-sm">{b.name || "—"}</p>
                    {b.phone && <p className="text-gray-500 text-xs">{b.phone}</p>}
                  </div>
                </div>
                <div className="text-right">
                  <p className="text-[#ffd025] font-black text-sm">${b.totalSpent.toLocaleString("es-CL")}</p>
                  <p className="text-gray-500 text-xs">{b.orderCount} {b.orderCount === 1 ? "pedido" : "pedidos"}</p>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Registered customers */}
      <div className="bg-[#1a1a1a] rounded-3xl border border-[#ffd025]/20 p-6">
        <h2 className="text-xl font-black uppercase mb-4 flex items-center gap-2">
          <Users size={20} className="text-[#ffd025]" /> Clientes registrados
          <span className="text-gray-500 text-sm font-normal">({customersData?.registered.length ?? 0})</span>
        </h2>
        {!customersData?.registered.length ? (
          <p className="text-gray-500 text-sm">Aún no hay clientes registrados.</p>
        ) : (
          <div className="space-y-2">
            {customersData.registered.map(c => (
              <div key={c.id} className="bg-[#141414] rounded-2xl px-4 py-3">
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0 flex-1">
                    <p className="font-bold text-sm truncate">{c.name}</p>
                    <p className="text-gray-400 text-xs truncate">{c.email}</p>
                    {c.phone && <p className="text-gray-500 text-xs">{c.phone}</p>}
                  </div>
                  <div className="text-right shrink-0">
                    <p className="text-gray-500 text-[10px]">Desde {new Date(c.createdAt).toLocaleDateString("es-CL")}</p>
                    <button
                      onClick={() => { setNewDiscount(p => ({ ...p, customerId: c.id })); setShowCreateForm(true); window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" }); }}
                      className="text-xs text-[#ffd025]/70 hover:text-[#ffd025] mt-1 transition-colors font-bold"
                    >
                      + Dar descuento
                    </button>
                  </div>
                </div>
                <div className="flex items-center gap-3 mt-2.5 pt-2.5 border-t border-gray-800">
                  <div className="flex items-center gap-1.5">
                    <ShoppingBag size={11} className="text-gray-500" />
                    <span className="text-xs text-gray-400">
                      {c.orderCount === 0
                        ? <span className="text-gray-600">Sin pedidos</span>
                        : <><span className="text-white font-bold">{c.orderCount}</span> {c.orderCount === 1 ? "pedido" : "pedidos"}</>}
                    </span>
                  </div>
                  {c.totalSpent > 0 && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-gray-600 text-[10px]">·</span>
                      <span className="text-xs text-[#ffd025] font-bold">${c.totalSpent.toLocaleString("es-CL")}</span>
                      <span className="text-xs text-gray-500">total</span>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Discount management */}
      <div className="bg-[#1a1a1a] rounded-3xl border border-[#ffd025]/20 p-6">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-xl font-black uppercase flex items-center gap-2">
            <Gift size={20} className="text-[#ffd025]" /> Descuentos
          </h2>
          <button
            onClick={() => { setNewDiscount({ code: "", type: "percentage", amount: 0, customerId: null, minOrder: 0, usesLeft: null, expiresAt: "" }); setShowCreateForm(!showCreateForm); }}
            className="flex items-center gap-2 px-4 py-2 bg-[#ffd025] text-[#141414] rounded-xl font-bold text-sm hover:bg-[#e5b81a] transition-colors"
          >
            <Plus size={16} /> Nuevo
          </button>
        </div>

        {showCreateForm && (
          <form onSubmit={createDiscount} className="bg-[#141414] rounded-2xl p-4 mb-5 space-y-3 border border-[#ffd025]/20">
            <h3 className="font-bold text-sm uppercase text-[#ffd025]">Crear código de descuento</h3>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">Código</label>
                <input type="text" value={newDiscount.code} onChange={e => setNewDiscount(p => ({ ...p, code: e.target.value.toUpperCase() }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm uppercase tracking-widest focus:outline-none focus:border-[#ffd025]" placeholder="DESCUENTO10" required />
              </div>
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">Tipo</label>
                <select value={newDiscount.type} onChange={e => setNewDiscount(p => ({ ...p, type: e.target.value as "percentage" | "fixed" }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]">
                  <option value="percentage">% Porcentaje</option>
                  <option value="fixed">$ Monto fijo</option>
                </select>
              </div>
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">{newDiscount.type === "percentage" ? "Porcentaje (%)" : "Monto ($)"}</label>
                <input type="number" min={1} max={newDiscount.type === "percentage" ? 100 : undefined} value={newDiscount.amount || ""} onChange={e => setNewDiscount(p => ({ ...p, amount: Number(e.target.value) }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="10" required />
              </div>
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">Mínimo de compra ($)</label>
                <input type="number" min={0} value={newDiscount.minOrder || ""} onChange={e => setNewDiscount(p => ({ ...p, minOrder: Number(e.target.value) }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="0" />
              </div>
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">Usos máx. (vacío = ∞)</label>
                <input type="number" min={1} value={newDiscount.usesLeft ?? ""} onChange={e => setNewDiscount(p => ({ ...p, usesLeft: e.target.value ? Number(e.target.value) : null }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]" placeholder="∞" />
              </div>
              <div>
                <label className="block text-[10px] text-gray-400 uppercase mb-1">Expira (opcional)</label>
                <input type="date" value={newDiscount.expiresAt} onChange={e => setNewDiscount(p => ({ ...p, expiresAt: e.target.value }))}
                  className="w-full bg-[#1a1a1a] border border-[#ffd025]/20 rounded-xl p-2.5 text-white text-sm focus:outline-none focus:border-[#ffd025]" />
              </div>
            </div>
            {newDiscount.customerId && (
              <div className="flex items-center gap-2 text-xs text-blue-400">
                <span>Descuento personal para cliente ID: {newDiscount.customerId}</span>
                <button type="button" onClick={() => setNewDiscount(p => ({ ...p, customerId: null }))} className="text-gray-500 hover:text-gray-300">✕ quitar</button>
              </div>
            )}
            <div className="flex gap-2 pt-1">
              <button type="submit" disabled={creating} className="flex-1 py-2.5 bg-[#ffd025] text-[#141414] rounded-xl font-bold text-sm hover:bg-[#e5b81a] transition-colors disabled:opacity-50">
                {creating ? "Creando..." : "Crear descuento"}
              </button>
              <button type="button" onClick={() => setShowCreateForm(false)} className="px-4 py-2.5 bg-[#0f0f0f] border border-gray-700 text-gray-400 rounded-xl font-bold text-sm hover:bg-gray-800 transition-colors">
                Cancelar
              </button>
            </div>
          </form>
        )}

        {discounts.length === 0 ? (
          <p className="text-gray-500 text-sm">Aún no hay códigos de descuento. Crea el primero.</p>
        ) : (
          <div className="space-y-2">
            {discounts.map(d => (
              <div key={d.id} className={`flex items-center justify-between rounded-2xl px-4 py-3 border ${d.active ? "bg-[#141414] border-[#ffd025]/10" : "bg-[#0f0f0f] border-gray-800 opacity-60"}`}>
                <div className="flex-1 min-w-0 mr-3">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-black text-sm tracking-widest text-[#ffd025]">{d.code}</span>
                    {d.customerId && <span className="text-[10px] bg-blue-900/40 text-blue-300 px-2 py-0.5 rounded-full border border-blue-800/40">Personal</span>}
                    {!d.active && <span className="text-[10px] bg-red-900/40 text-red-300 px-2 py-0.5 rounded-full border border-red-800/40">Inactivo</span>}
                  </div>
                  <p className="text-gray-400 text-xs mt-0.5 truncate">
                    {d.type === "percentage" ? `${d.amount}% desc.` : `$${d.amount.toLocaleString("es-CL")} desc.`}
                    {d.minOrder > 0 && ` · mín. $${d.minOrder.toLocaleString("es-CL")}`}
                    {d.usesLeft !== null && ` · ${d.usesLeft} usos restantes`}
                    {d.customerName && ` · para ${d.customerName}`}
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button onClick={() => toggleDiscount(d.id, !d.active)}
                    className={`text-xs px-3 py-1.5 rounded-xl font-bold transition-colors ${d.active ? "bg-green-900/30 text-green-400 hover:bg-green-900/50" : "bg-gray-800 text-gray-400 hover:bg-gray-700"}`}>
                    {d.active ? "Activo" : "Inactivo"}
                  </button>
                  <button onClick={() => deleteDiscount(d.id)} className="text-red-400/50 hover:text-red-400 p-1.5 rounded-xl hover:bg-red-900/20 transition-colors">
                    <Trash2 size={14} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function StatCard({
  label,
  value,
  icon,
}: {
  label: string;
  value: number | string;
  icon: React.ReactNode;
}) {
  return (
    <div className="bg-[#141414] border border-gray-800 rounded-2xl p-4">
      <div className="flex items-center gap-2 text-gray-500 text-[10px] uppercase font-bold tracking-wider mb-1">
        <span className="text-[#ffd025]">{icon}</span>
        {label}
      </div>
      <div className="text-2xl font-black text-[#ffd025]">{value}</div>
    </div>
  );
}
