import fs from "fs";
import path from "path";

const DATA_DIR = path.resolve(process.cwd(), "data");
const DB_FILE = path.join(DATA_DIR, "fellas_db.json");
const UPLOADS_DIR = path.join(DATA_DIR, "uploads");

// Ensure directories exist
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

export interface Setting {
  id?: number;
  logo: string;
  pageTitle: string;
  favicon: string;
  bannerText: string;
  bannerDescription: string;
  bannerImage: string;
  bannerSlides: Array<{ image: string; title: string; description: string }>;
  whatsapp: string;
  agencyName: string;
  agencyLogo: string;
  announcements: string[];
  contactPhone: string;
  contactAddress: string;
  contactHours: string;
  openTime: string;
  closeTime: string;
  deliveryMinimum: number;
}

export interface Category {
  id: number;
  name: string;
  position: number;
}

export interface Aisle {
  id: number;
  name: string;
  bannerImage: string;
  bannerTitle: string;
  bannerSubtitle: string;
  position?: number;
}

export interface Subcategory {
  id: number;
  name: string;
  position?: number;
}

export interface Product {
  id: number;
  name: string;
  price: number;
  image: string;
  category: string;
  aisle: string;
  subcategory: string;
  optionsTitle: string;
  options: string[];
  bestseller?: boolean;
  oferta: boolean;
  depositoEnabled: boolean;
  depositoAmount: number;
  position: number;
  publishedSocial: boolean;
  hidden: boolean;
  transferenciaEnabled: boolean;
  transferenciaAmount: number;
}

export interface OrderItem {
  name: string;
  quantity: number;
  price: number;
  selectedOption?: string;
}

export interface Order {
  id: number;
  customerName: string;
  address: string;
  phone: string;
  notes: string;
  items: OrderItem[];
  subtotal?: number | null;
  discountCode?: string | null;
  discountAmount?: number | null;
  total: number;
  status: string;
  caja?: string | null;
  modificationNote?: string | null;
  createdAt: string;
}

export interface DeliveryLocation {
  id: number;
  name: string;
  price: number;
  active: boolean;
  position: number;
}

export interface Discount {
  id: number;
  code: string;
  type: string;
  amount: number;
  customerId?: number | null;
  active: boolean;
  minOrder: number;
  usesLeft?: number | null;
  expiresAt?: string | null;
  createdAt: string;
}

export interface Customer {
  id: number;
  email: string;
  passwordHash: string;
  name: string;
  phone: string;
  createdAt: string;
}

export interface SavedCart {
  id: number;
  customerId: number;
  name: string;
  items: OrderItem[];
  createdAt: string;
}

export interface Visit {
  id: number;
  ipAddress: string;
  city: string;
  region: string;
  country: string;
  path: string;
  userAgent: string;
  createdAt: string;
}

export interface DatabaseSchema {
  settings: Setting;
  categories: Category[];
  aisles: Aisle[];
  subcategories: Subcategory[];
  products: Product[];
  orders: Order[];
  deliveryLocations: DeliveryLocation[];
  discounts: Discount[];
  customers: Customer[];
  savedCarts: SavedCart[];
  visits: Visit[];
}

const DEFAULT_DB: DatabaseSchema = {
  settings: {
    logo: "",
    pageTitle: "Fella's Market — Botillería en Alerce, Puerto Montt",
    favicon: "/favicon.svg",
    bannerText: "🍻 ¡BIENVENIDO A FELLA'S MARKET ALERCE! 🍾",
    bannerDescription: "Tu botillería y minimarket de confianza. Cervezas heladas, destilados, vinos, snacks y delivery rápido a todo Alerce y Puerto Montt.",
    bannerImage: "/src/assets/banner.png",
    bannerSlides: [
      {
        image: "/src/assets/banner.png",
        title: "FELLA'S MARKET ALERCE",
        description: "Bebidas heladas, licores y botillería completa con despacho a domicilio."
      },
      {
        image: "/src/assets/burger.png",
        title: "PROMO PISCOLA DEL FIN DE SEMANA",
        description: "Mistral o Alto del Carmen + Bebida 1.5L + Hielo 2kg a precio especial."
      },
      {
        image: "/src/assets/drink.png",
        title: "CERVEZAS NACIONALES E IMPORTADAS",
        description: "Packs de Corona, Heineken, Kunstmann y Austral listas para disfrutar."
      }
    ],
    whatsapp: "56912345678",
    agencyName: "Fella's Market",
    agencyLogo: "",
    announcements: [
      "🛵 Delivery activo en Alerce Histórico, Alerce Norte y Alerce Sur",
      "❄️ Hielo en bolsa 2kg y 4kg siempre disponible",
      "💳 Transferencia bancaria y efectivo aceptados"
    ],
    contactPhone: "+56 9 1234 5678",
    contactAddress: "Alerce, Puerto Montt, Chile",
    contactHours: "Lunes a Domingo: 12:00 a 03:00 hrs",
    openTime: "00:00",
    closeTime: "23:59",
    deliveryMinimum: 5000,
  },
  categories: [
    { id: 1, name: "Cervezas", position: 0 },
    { id: 2, name: "Pisco & Destilados", position: 1 },
    { id: 3, name: "Vinos & Espumantes", position: 2 },
    { id: 4, name: "Bebidas & Energéticas", position: 3 },
    { id: 5, name: "Snacks & Picoteo", position: 4 },
    { id: 6, name: "Hielo & Promos", position: 5 },
  ],
  aisles: [
    {
      id: 1,
      name: "Botillería",
      bannerImage: "/src/assets/drink.png",
      bannerTitle: "Botillería Fella's",
      bannerSubtitle: "Pisco, cervezas heladas, whisky, ron y vodkas",
      position: 0,
    },
    {
      id: 2,
      name: "Minimarket",
      bannerImage: "/src/assets/fries.png",
      bannerTitle: "Minimarket & Picoteo",
      bannerSubtitle: "Papas fritas, galletas, chocolates y bebidas desechables",
      position: 1,
    },
    {
      id: 3,
      name: "Promociones",
      bannerImage: "/src/assets/banner.png",
      bannerTitle: "Combos & Ofertas",
      bannerSubtitle: "Packs fiesta armados para disfrutar al máximo",
      position: 2,
    },
  ],
  subcategories: [
    { id: 1, name: "Cervezas Nacionales", position: 0 },
    { id: 2, name: "Cervezas Artesanales", position: 1 },
    { id: 3, name: "Pisco Chileno", position: 2 },
    { id: 4, name: "Whisky & Gin", position: 3 },
    { id: 5, name: "Bebidas Gaseosas", position: 4 },
    { id: 6, name: "Snacks Salados", position: 5 },
  ],
  products: [
    {
      id: 1,
      name: "Pack Promo Piscola: Mistral 35° + Coca-Cola 1.5L + Hielo 2kg",
      price: 10990,
      image: "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=600&auto=format&fit=crop&q=80",
      category: "Hielo & Promos",
      aisle: "Promociones",
      subcategory: "Pisco Chileno",
      optionsTitle: "Variedad de Bebida",
      options: ["Coca-Cola Normal", "Coca-Cola Zero"],
      oferta: true,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 0,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 2,
      name: "Pisco Mistral 35° 750ml",
      price: 6990,
      image: "https://images.unsplash.com/photo-1527061011665-3652c757a4d4?w=600&auto=format&fit=crop&q=80",
      category: "Pisco & Destilados",
      aisle: "Botillería",
      subcategory: "Pisco Chileno",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 1,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 3,
      name: "Pisco Alto del Carmen 35° Especial 1 Litro",
      price: 8490,
      image: "https://images.unsplash.com/photo-1569529465841-dfecdab7503b?w=600&auto=format&fit=crop&q=80",
      category: "Pisco & Destilados",
      aisle: "Botillería",
      subcategory: "Pisco Chileno",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 2,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 4,
      name: "Pack 6 Cervezas Corona Extra 330ml",
      price: 6990,
      image: "https://images.unsplash.com/photo-1608270191763-7182586df111?w=600&auto=format&fit=crop&q=80",
      category: "Cervezas",
      aisle: "Botillería",
      subcategory: "Cervezas Nacionales",
      optionsTitle: "",
      options: [],
      oferta: true,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 3,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 5,
      name: "Pack 6 Cervezas Heineken 330ml Botella",
      price: 7490,
      image: "https://images.unsplash.com/photo-1618886614638-80e3c15cd819?w=600&auto=format&fit=crop&q=80",
      category: "Cervezas",
      aisle: "Botillería",
      subcategory: "Cervezas Nacionales",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 4,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 6,
      name: "Cerveza Kunstmann Torobayo 330ml",
      price: 1990,
      image: "https://images.unsplash.com/photo-1571613316887-6f8d5cbf7ef7?w=600&auto=format&fit=crop&q=80",
      category: "Cervezas",
      aisle: "Botillería",
      subcategory: "Cervezas Artesanales",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 5,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 7,
      name: "Cerveza Austral Calafate 330ml",
      price: 2190,
      image: "https://images.unsplash.com/photo-1535958636474-b021ee887b13?w=600&auto=format&fit=crop&q=80",
      category: "Cervezas",
      aisle: "Botillería",
      subcategory: "Cervezas Artesanales",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 6,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 8,
      name: "Whisky Johnnie Walker Red Label 750ml",
      price: 14990,
      image: "https://images.unsplash.com/photo-1527281400683-1aae777175f8?w=600&auto=format&fit=crop&q=80",
      category: "Pisco & Destilados",
      aisle: "Botillería",
      subcategory: "Whisky & Gin",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 7,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 9,
      name: "Gin Beefeater London Dry 750ml",
      price: 15990,
      image: "https://images.unsplash.com/photo-1514362545857-3bc16c4c7d1b?w=600&auto=format&fit=crop&q=80",
      category: "Pisco & Destilados",
      aisle: "Botillería",
      subcategory: "Whisky & Gin",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 8,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 10,
      name: "Coca-Cola Original 1.5L Desechable",
      price: 1890,
      image: "https://images.unsplash.com/photo-1622483767028-3f66f32aef97?w=600&auto=format&fit=crop&q=80",
      category: "Bebidas & Energéticas",
      aisle: "Minimarket",
      subcategory: "Bebidas Gaseosas",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 9,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 11,
      name: "Bebida Energética Monster Energy 473ml",
      price: 2200,
      image: "https://images.unsplash.com/photo-1622543925917-763c34d1a86e?w=600&auto=format&fit=crop&q=80",
      category: "Bebidas & Energéticas",
      aisle: "Minimarket",
      subcategory: "Bebidas Gaseosas",
      optionsTitle: "Sabor",
      options: ["Original Verde", "Mango Loco", "Ultra White Zero"],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 10,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 12,
      name: "Vino Casillero del Diablo Cabernet Sauvignon 750ml",
      price: 5490,
      image: "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=600&auto=format&fit=crop&q=80",
      category: "Vinos & Espumantes",
      aisle: "Botillería",
      subcategory: "Pisco Chileno",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 11,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 13,
      name: "Espumante Valdivieso Brut 750ml",
      price: 5990,
      image: "https://images.unsplash.com/photo-1568213816046-0ee1c42bd559?w=600&auto=format&fit=crop&q=80",
      category: "Vinos & Espumantes",
      aisle: "Botillería",
      subcategory: "Pisco Chileno",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 12,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 14,
      name: "Papas Fritas Lays Clásicas 250g",
      price: 2490,
      image: "https://images.unsplash.com/photo-1566478989037-eec170784d0b?w=600&auto=format&fit=crop&q=80",
      category: "Snacks & Picoteo",
      aisle: "Minimarket",
      subcategory: "Snacks Salados",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 13,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 15,
      name: "Doritos Mega Queso 200g",
      price: 2290,
      image: "https://images.unsplash.com/photo-1584985474011-8a62372e9a8f?w=600&auto=format&fit=crop&q=80",
      category: "Snacks & Picoteo",
      aisle: "Minimarket",
      subcategory: "Snacks Salados",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 14,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 16,
      name: "Bolsa de Hielo Purificado en Cubos 2kg",
      price: 1500,
      image: "https://images.unsplash.com/photo-1544816155-12df9643f363?w=600&auto=format&fit=crop&q=80",
      category: "Hielo & Promos",
      aisle: "Botillería",
      subcategory: "Cervezas Nacionales",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 15,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
    {
      id: 17,
      name: "Bolsa de Hielo Purificado en Cubos 4kg",
      price: 2500,
      image: "https://images.unsplash.com/photo-1544816155-12df9643f363?w=600&auto=format&fit=crop&q=80",
      category: "Hielo & Promos",
      aisle: "Botillería",
      subcategory: "Cervezas Nacionales",
      optionsTitle: "",
      options: [],
      oferta: false,
      depositoEnabled: false,
      depositoAmount: 0,
      position: 16,
      publishedSocial: true,
      hidden: false,
      transferenciaEnabled: false,
      transferenciaAmount: 0,
    },
  ],
  orders: [
    {
      id: 101,
      customerName: "Matías Almonacid",
      address: "Pasaje Los Notros 142, Alerce Histórico",
      phone: "+56987654321",
      notes: "Por favor tocar el timbre blanco.",
      items: [
        { name: "Pack Promo Piscola: Mistral 35° + Coca-Cola 1.5L + Hielo 2kg", quantity: 1, price: 10990, selectedOption: "Coca-Cola Normal" },
        { name: "Papas Fritas Lays Clásicas 250g", quantity: 1, price: 2490 }
      ],
      subtotal: 13480,
      discountCode: null,
      discountAmount: null,
      total: 13480,
      status: "Entregado",
      caja: "Caja 1",
      modificationNote: null,
      createdAt: new Date(Date.now() - 3600 * 1000 * 4).toISOString(),
    },
    {
      id: 102,
      customerName: "Camila Soto",
      address: "Av. Gabriela Mistral 850, Alerce Norte",
      phone: "+56976543210",
      notes: "Llamar al llegar",
      items: [
        { name: "Pack 6 Cervezas Corona Extra 330ml", quantity: 2, price: 6990 },
        { name: "Bolsa de Hielo Purificado en Cubos 2kg", quantity: 1, price: 1500 }
      ],
      subtotal: 15480,
      discountCode: null,
      discountAmount: null,
      total: 15480,
      status: "Confirmado y en Preparación",
      caja: "Caja 1",
      modificationNote: null,
      createdAt: new Date(Date.now() - 3600 * 1000 * 1).toISOString(),
    }
  ],
  deliveryLocations: [
    { id: 1, name: "Alerce Histórico", price: 1500, active: true, position: 0 },
    { id: 2, name: "Alerce Norte", price: 2000, active: true, position: 1 },
    { id: 3, name: "Alerce Sur", price: 2000, active: true, position: 2 },
    { id: 4, name: "Senda Central / Polincay", price: 2500, active: true, position: 3 },
    { id: 5, name: "Bosquemar / La Vara", price: 3000, active: true, position: 4 },
    { id: 6, name: "Retiro en Local Fella's Market", price: 0, active: true, position: 5 },
  ],
  discounts: [
    {
      id: 1,
      code: "FELLAS10",
      type: "percentage",
      amount: 10,
      customerId: null,
      active: true,
      minOrder: 10000,
      usesLeft: 100,
      expiresAt: null,
      createdAt: new Date().toISOString(),
    },
    {
      id: 2,
      code: "ALERCEFIESTA",
      type: "fixed",
      amount: 2000,
      customerId: null,
      active: true,
      minOrder: 15000,
      usesLeft: 50,
      expiresAt: null,
      createdAt: new Date().toISOString(),
    }
  ],
  customers: [],
  savedCarts: [],
  visits: [
    {
      id: 1,
      ipAddress: "190.161.45.12",
      city: "Puerto Montt",
      region: "Los Lagos",
      country: "Chile",
      path: "/",
      userAgent: "Mozilla/5.0 (iPhone)",
      createdAt: new Date(Date.now() - 3600 * 1000 * 2).toISOString(),
    },
    {
      id: 2,
      ipAddress: "186.104.22.8",
      city: "Puerto Montt",
      region: "Los Lagos",
      country: "Chile",
      path: "/",
      userAgent: "Mozilla/5.0 (Android)",
      createdAt: new Date(Date.now() - 3600 * 1000 * 5).toISOString(),
    },
    {
      id: 3,
      ipAddress: "201.246.10.99",
      city: "Puerto Varas",
      region: "Los Lagos",
      country: "Chile",
      path: "/",
      userAgent: "Mozilla/5.0 (Windows)",
      createdAt: new Date(Date.now() - 3600 * 1000 * 12).toISOString(),
    }
  ],
};

class DatabaseManager {
  private db: DatabaseSchema;
  private saveTimeout: NodeJS.Timeout | null = null;

  constructor() {
    this.db = this.load();
  }

  private load(): DatabaseSchema {
    try {
      if (fs.existsSync(DB_FILE)) {
        const raw = fs.readFileSync(DB_FILE, "utf-8");
        const parsed = JSON.parse(raw);
        // Merge defaults in case new fields were added
        return {
          ...DEFAULT_DB,
          ...parsed,
          settings: { ...DEFAULT_DB.settings, ...(parsed.settings || {}) },
          categories: parsed.categories || DEFAULT_DB.categories,
          aisles: parsed.aisles || DEFAULT_DB.aisles,
          subcategories: parsed.subcategories || DEFAULT_DB.subcategories,
          products: parsed.products || DEFAULT_DB.products,
          orders: parsed.orders || DEFAULT_DB.orders,
          deliveryLocations: parsed.deliveryLocations || DEFAULT_DB.deliveryLocations,
          discounts: parsed.discounts || DEFAULT_DB.discounts,
          customers: parsed.customers || DEFAULT_DB.customers,
          savedCarts: parsed.savedCarts || DEFAULT_DB.savedCarts,
          visits: parsed.visits || DEFAULT_DB.visits,
        };
      }
    } catch (err) {
      console.error("Error reading database file, resetting to defaults:", err);
    }
    this.saveImmediate(DEFAULT_DB);
    return JSON.parse(JSON.stringify(DEFAULT_DB));
  }

  private saveImmediate(data: DatabaseSchema) {
    try {
      const tmp = `${DB_FILE}.tmp`;
      fs.writeFileSync(tmp, JSON.stringify(data, null, 2), "utf-8");
      fs.renameSync(tmp, DB_FILE);
    } catch (err) {
      console.error("Failed to write database file:", err);
    }
  }

  public save() {
    if (this.saveTimeout) clearTimeout(this.saveTimeout);
    this.saveTimeout = setTimeout(() => {
      this.saveImmediate(this.db);
    }, 100);
  }

  public getSettings(): Setting {
    return this.db.settings;
  }

  public updateSettings(partial: Partial<Setting>): Setting {
    this.db.settings = { ...this.db.settings, ...partial };
    this.save();
    return this.db.settings;
  }

  public getCategories(): Category[] {
    return [...this.db.categories].sort((a, b) => (a.position ?? 0) - (b.position ?? 0) || a.id - b.id);
  }

  public addCategory(name: string): Category {
    const nextId = this.db.categories.length ? Math.max(...this.db.categories.map(c => c.id)) + 1 : 1;
    const cat: Category = { id: nextId, name, position: this.db.categories.length };
    this.db.categories.push(cat);
    this.save();
    return cat;
  }

  public deleteCategory(id: number): boolean {
    const idx = this.db.categories.findIndex(c => c.id === id);
    if (idx !== -1) {
      this.db.categories.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public reorderCategories(ids: number[]) {
    ids.forEach((id, idx) => {
      const cat = this.db.categories.find(c => c.id === id);
      if (cat) cat.position = idx;
    });
    this.save();
  }

  public getAisles(): Aisle[] {
    return [...this.db.aisles].sort((a, b) => (a.position ?? 0) - (b.position ?? 0) || a.id - b.id);
  }

  public addAisle(name: string): Aisle {
    const nextId = this.db.aisles.length ? Math.max(...this.db.aisles.map(a => a.id)) + 1 : 1;
    const aisle: Aisle = {
      id: nextId,
      name,
      bannerImage: "",
      bannerTitle: name,
      bannerSubtitle: "",
      position: this.db.aisles.length,
    };
    this.db.aisles.push(aisle);
    this.save();
    return aisle;
  }

  public updateAisle(id: number, partial: Partial<Aisle>): Aisle | null {
    const aisle = this.db.aisles.find(a => a.id === id);
    if (!aisle) return null;
    Object.assign(aisle, partial);
    this.save();
    return aisle;
  }

  public deleteAisle(id: number): boolean {
    const idx = this.db.aisles.findIndex(a => a.id === id);
    if (idx !== -1) {
      this.db.aisles.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public reorderAisles(ids: number[]) {
    ids.forEach((id, idx) => {
      const aisle = this.db.aisles.find(a => a.id === id);
      if (aisle) aisle.position = idx;
    });
    this.save();
  }

  public getSubcategories(): Subcategory[] {
    return [...this.db.subcategories].sort((a, b) => (a.position ?? 0) - (b.position ?? 0) || a.id - b.id);
  }

  public addSubcategory(name: string): Subcategory {
    const nextId = this.db.subcategories.length ? Math.max(...this.db.subcategories.map(s => s.id)) + 1 : 1;
    const sub: Subcategory = { id: nextId, name, position: this.db.subcategories.length };
    this.db.subcategories.push(sub);
    this.save();
    return sub;
  }

  public deleteSubcategory(id: number): boolean {
    const idx = this.db.subcategories.findIndex(s => s.id === id);
    if (idx !== -1) {
      this.db.subcategories.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public reorderSubcategories(ids: number[]) {
    ids.forEach((id, idx) => {
      const sub = this.db.subcategories.find(s => s.id === id);
      if (sub) sub.position = idx;
    });
    this.save();
  }

  public getProducts(): Product[] {
    return [...this.db.products].sort((a, b) => (a.position ?? 0) - (b.position ?? 0) || a.id - b.id);
  }

  public addProduct(p: Omit<Product, "id">): Product {
    const nextId = this.db.products.length ? Math.max(...this.db.products.map(x => x.id)) + 1 : 1;
    const prod: Product = { ...p, id: nextId };
    this.db.products.push(prod);
    this.save();
    return prod;
  }

  public updateProduct(id: number, partial: Partial<Product>): Product | null {
    const prod = this.db.products.find(p => p.id === id);
    if (!prod) return null;
    Object.assign(prod, partial);
    this.save();
    return prod;
  }

  public deleteProduct(id: number): boolean {
    const idx = this.db.products.findIndex(p => p.id === id);
    if (idx !== -1) {
      this.db.products.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public reorderProducts(ids: number[]) {
    ids.forEach((id, idx) => {
      const prod = this.db.products.find(p => p.id === id);
      if (prod) prod.position = idx;
    });
    this.save();
  }

  public getOrders(): Order[] {
    return [...this.db.orders].sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  }

  public addOrder(o: Omit<Order, "id" | "createdAt">): Order {
    const nextId = this.db.orders.length ? Math.max(...this.db.orders.map(x => x.id)) + 1 : 101;
    const order: Order = {
      ...o,
      id: nextId,
      createdAt: new Date().toISOString(),
    };
    this.db.orders.push(order);
    this.save();
    return order;
  }

  public updateOrder(id: number, partial: Partial<Order>): Order | null {
    const order = this.db.orders.find(o => o.id === id);
    if (!order) return null;
    Object.assign(order, partial);
    this.save();
    return order;
  }

  public deleteOrder(id: number): boolean {
    const idx = this.db.orders.findIndex(o => o.id === id);
    if (idx !== -1) {
      this.db.orders.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public getDeliveryLocations(): DeliveryLocation[] {
    return [...this.db.deliveryLocations].sort((a, b) => (a.position ?? 0) - (b.position ?? 0) || a.id - b.id);
  }

  public addDeliveryLocation(loc: Omit<DeliveryLocation, "id">): DeliveryLocation {
    const nextId = this.db.deliveryLocations.length ? Math.max(...this.db.deliveryLocations.map(l => l.id)) + 1 : 1;
    const created: DeliveryLocation = { ...loc, id: nextId };
    this.db.deliveryLocations.push(created);
    this.save();
    return created;
  }

  public updateDeliveryLocation(id: number, partial: Partial<DeliveryLocation>): DeliveryLocation | null {
    const loc = this.db.deliveryLocations.find(l => l.id === id);
    if (!loc) return null;
    Object.assign(loc, partial);
    this.save();
    return loc;
  }

  public deleteDeliveryLocation(id: number): boolean {
    const idx = this.db.deliveryLocations.findIndex(l => l.id === id);
    if (idx !== -1) {
      this.db.deliveryLocations.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public getDiscounts(): Discount[] {
    return this.db.discounts;
  }

  public getDiscountByCode(code: string): Discount | null {
    return this.db.discounts.find(d => d.code.toUpperCase() === code.trim().toUpperCase() && d.active) || null;
  }

  public addDiscount(d: Omit<Discount, "id" | "createdAt">): Discount {
    const nextId = this.db.discounts.length ? Math.max(...this.db.discounts.map(x => x.id)) + 1 : 1;
    const created: Discount = { ...d, id: nextId, createdAt: new Date().toISOString() };
    this.db.discounts.push(created);
    this.save();
    return created;
  }

  public deleteDiscount(id: number): boolean {
    const idx = this.db.discounts.findIndex(d => d.id === id);
    if (idx !== -1) {
      this.db.discounts.splice(idx, 1);
      this.save();
      return true;
    }
    return false;
  }

  public addVisit(v: Omit<Visit, "id" | "createdAt">): Visit {
    const nextId = this.db.visits.length ? Math.max(...this.db.visits.map(x => x.id)) + 1 : 1;
    const visit: Visit = { ...v, id: nextId, createdAt: new Date().toISOString() };
    this.db.visits.push(visit);
    this.save();
    return visit;
  }

  public getVisits(): Visit[] {
    return this.db.visits;
  }

  public getCustomerByEmail(email: string): Customer | null {
    return this.db.customers.find(c => c.email.toLowerCase() === email.toLowerCase().trim()) || null;
  }

  public addCustomer(c: Omit<Customer, "id" | "createdAt">): Customer {
    const nextId = this.db.customers.length ? Math.max(...this.db.customers.map(x => x.id)) + 1 : 1;
    const customer: Customer = { ...c, id: nextId, createdAt: new Date().toISOString() };
    this.db.customers.push(customer);
    this.save();
    return customer;
  }

  public getCustomerById(id: number): Customer | null {
    return this.db.customers.find(c => c.id === id) || null;
  }

  public getSavedCart(customerId: number): SavedCart | null {
    return this.db.savedCarts.find(s => s.customerId === customerId) || null;
  }

  public saveCart(customerId: number, name: string, items: OrderItem[]): SavedCart {
    let cart = this.db.savedCarts.find(s => s.customerId === customerId);
    if (cart) {
      cart.items = items;
      cart.name = name;
      cart.createdAt = new Date().toISOString();
    } else {
      const nextId = this.db.savedCarts.length ? Math.max(...this.db.savedCarts.map(x => x.id)) + 1 : 1;
      cart = { id: nextId, customerId, name, items, createdAt: new Date().toISOString() };
      this.db.savedCarts.push(cart);
    }
    this.save();
    return cart;
  }
}

export const dbManager = new DatabaseManager();
