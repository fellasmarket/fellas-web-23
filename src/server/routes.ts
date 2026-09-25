import express, { type Request, type Response } from "express";
import fs from "fs";
import path from "path";
import { randomUUID } from "crypto";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { GoogleGenAI } from "@google/genai";
import { dbManager, type Product, type Order, type OrderItem } from "./db.ts";

const JWT_SECRET = process.env.JWT_SECRET || "fellas-market-secret-jwt-key-2024";
const DATA_DIR = path.resolve(process.cwd(), "data");
const UPLOADS_DIR = path.join(DATA_DIR, "uploads");

if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

export const apiRouter = express.Router();

// In-memory cache for ip -> geo
const ipGeoCache = new Map<string, { city: string; region: string; country: string }>();

async function lookupIpGeo(ip: string) {
  if (
    !ip ||
    ip === "::1" ||
    ip.startsWith("127.") ||
    ip.startsWith("10.") ||
    ip.startsWith("192.168.") ||
    ip.startsWith("172.")
  ) {
    return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
  }
  const cached = ipGeoCache.get(ip);
  if (cached) return cached;
  try {
    const res = await fetch(
      `http://ip-api.com/json/${encodeURIComponent(ip)}?fields=status,country,regionName,city`,
      { signal: AbortSignal.timeout(2500) }
    );
    if (!res.ok) return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
    const data = (await res.json()) as { status?: string; country?: string; regionName?: string; city?: string };
    if (data.status !== "success") return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
    const geo = {
      city: data.city || "Puerto Montt",
      region: data.regionName || "Los Lagos",
      country: data.country || "Chile",
    };
    ipGeoCache.set(ip, geo);
    return geo;
  } catch {
    return { city: "Puerto Montt", region: "Los Lagos", country: "Chile" };
  }
}

function getCustomerIdFromReq(req: Request): number | null {
  const auth = req.headers.authorization;
  if (!auth?.startsWith("Bearer ")) return null;
  try {
    const token = auth.slice(7);
    const decoded = jwt.verify(token, JWT_SECRET) as { customerId: number };
    return decoded.customerId ?? null;
  } catch {
    return null;
  }
}

// 1. Health
apiRouter.get("/healthz", (_req, res) => {
  res.json({ status: "ok" });
});

// 2. Menu
apiRouter.get("/menu", (_req, res) => {
  const settings = dbManager.getSettings();
  const categories = dbManager.getCategories();
  const aisles = dbManager.getAisles();
  const subcategories = dbManager.getSubcategories();
  const allProducts = dbManager.getProducts();
  const orders = dbManager.getOrders();

  // Compute bestsellers from last 30 days orders
  const thirtyDaysAgo = Date.now() - 30 * 24 * 3600 * 1000;
  const productQtyMap = new Map<string, number>();

  for (const order of orders) {
    if (order.status === "Anulado") continue;
    if (new Date(order.createdAt).getTime() < thirtyDaysAgo) continue;
    for (const it of order.items || []) {
      const q = Number(it.quantity) || 1;
      productQtyMap.set(it.name, (productQtyMap.get(it.name) || 0) + q);
    }
  }

  const topNames = new Set(
    [...productQtyMap.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name]) => name)
  );

  const products = allProducts
    .filter((p) => !p.hidden)
    .map((p) => ({
      ...p,
      bestseller: topNames.has(p.name) || p.oferta,
    }));

  res.json({
    settings,
    categories,
    aisles,
    subcategories,
    products,
  });
});

// 3. Settings
apiRouter.get("/settings", (_req, res) => {
  res.json(dbManager.getSettings());
});

apiRouter.patch("/settings", (req, res) => {
  const updated = dbManager.updateSettings(req.body);
  res.json(updated);
});

// 4. Categories
apiRouter.get("/categories", (_req, res) => {
  res.json(dbManager.getCategories());
});

apiRouter.post("/categories", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addCategory(name);
  res.status(201).json(created);
});

apiRouter.delete("/categories/:id", (req, res) => {
  const ok = dbManager.deleteCategory(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/categories/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderCategories(ids);
  res.json({ ok: true });
});

// 5. Aisles
apiRouter.get("/aisles", (_req, res) => {
  res.json(dbManager.getAisles());
});

apiRouter.post("/aisles", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addAisle(name);
  res.status(201).json(created);
});

apiRouter.patch("/aisles/:id", (req, res) => {
  const updated = dbManager.updateAisle(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Aisle not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/aisles/:id", (req, res) => {
  const ok = dbManager.deleteAisle(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/aisles/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderAisles(ids);
  res.json({ ok: true });
});

// 6. Subcategories
apiRouter.get("/subcategories", (_req, res) => {
  res.json(dbManager.getSubcategories());
});

apiRouter.post("/subcategories", (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) {
    res.status(400).json({ error: "Name is required" });
    return;
  }
  const created = dbManager.addSubcategory(name);
  res.status(201).json(created);
});

apiRouter.delete("/subcategories/:id", (req, res) => {
  const ok = dbManager.deleteSubcategory(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/subcategories/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderSubcategories(ids);
  res.json({ ok: true });
});

// 7. Products
apiRouter.get("/products", (_req, res) => {
  res.json(dbManager.getProducts());
});

apiRouter.post("/products", (req, res) => {
  const body = req.body;
  const created = dbManager.addProduct({
    name: body.name,
    price: Number(body.price) || 0,
    image: body.image || "",
    category: body.category || "",
    aisle: body.aisle || "",
    subcategory: body.subcategory || "",
    optionsTitle: body.optionsTitle || "",
    options: Array.isArray(body.options) ? body.options : [],
    oferta: Boolean(body.oferta),
    depositoEnabled: Boolean(body.depositoEnabled),
    depositoAmount: Number(body.depositoAmount) || 0,
    position: dbManager.getProducts().length,
    publishedSocial: Boolean(body.publishedSocial),
    hidden: Boolean(body.hidden),
    transferenciaEnabled: Boolean(body.transferenciaEnabled),
    transferenciaAmount: Number(body.transferenciaAmount) || 0,
  });
  res.status(201).json(created);
});

apiRouter.patch("/products/:id", (req, res) => {
  const updated = dbManager.updateProduct(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Product not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/products/:id", (req, res) => {
  const ok = dbManager.deleteProduct(Number(req.params.id));
  res.json({ ok });
});

apiRouter.post("/products/reorder", (req, res) => {
  const ids: number[] = req.body.ids || [];
  dbManager.reorderProducts(ids);
  res.json({ ok: true });
});

// 8. Orders
apiRouter.get("/orders", (_req, res) => {
  res.json(dbManager.getOrders());
});

apiRouter.post("/orders", (req, res) => {
  const body = req.body;
  const order = dbManager.addOrder({
    customerName: body.customerName,
    address: body.address || "",
    phone: body.phone || "",
    notes: body.notes || "",
    items: body.items || [],
    subtotal: body.subtotal ?? null,
    discountCode: body.discountCode ?? null,
    discountAmount: body.discountAmount ?? null,
    total: Number(body.total) || 0,
    status: body.status || "Pendiente",
    caja: body.caja ?? null,
    modificationNote: body.modificationNote ?? null,
  });
  res.status(201).json(order);
});

apiRouter.patch("/orders/:id", (req, res) => {
  const id = Number(req.params.id);
  const updated = dbManager.updateOrder(id, req.body);
  if (!updated) {
    res.status(404).json({ error: "Order not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/orders/:id", (req, res) => {
  const ok = dbManager.deleteOrder(Number(req.params.id));
  res.json({ ok });
});

// 9. Admin
apiRouter.post("/admin/login", (req, res) => {
  const password = String(req.body.password || "").trim();
  const configuredPassword = process.env.ADMIN_PASSWORD || "fellhonpm";
  const deliveryPassword = process.env.DELIVERY_PASSWORD || "botifelldely";

  if (password === "fellhonpm" || password === configuredPassword) {
    res.json({ ok: true, role: "admin" });
    return;
  }
  if (password === "botifelldely" || password === deliveryPassword) {
    res.json({ ok: true, role: "delivery" });
    return;
  }
  res.status(401).json({ ok: false, error: "Contraseña incorrecta" });
});

apiRouter.get("/admin/stats", (_req, res) => {
  res.json({
    productCount: dbManager.getProducts().length,
    categoryCount: dbManager.getCategories().length,
    aisleCount: dbManager.getAisles().length,
    subcategoryCount: dbManager.getSubcategories().length,
  });
});

apiRouter.get("/admin/visits/stats", (_req, res) => {
  const visits = dbManager.getVisits();
  const now = Date.now();
  const dayMs = 24 * 3600 * 1000;
  const startOfToday = new Date().setHours(0, 0, 0, 0);

  const todayCount = visits.filter((v) => new Date(v.createdAt).getTime() >= startOfToday).length;
  const last7Count = visits.filter((v) => new Date(v.createdAt).getTime() >= now - 7 * dayMs).length;
  const last30Count = visits.filter((v) => new Date(v.createdAt).getTime() >= now - 30 * dayMs).length;

  const hourMap = new Map<number, number>();
  for (let h = 0; h < 24; h++) hourMap.set(h, 0);

  const cityMap = new Map<string, { city: string; country: string; count: number }>();
  const countryMap = new Map<string, number>();
  const dayMap = new Map<string, number>();

  for (const v of visits) {
    const vDate = new Date(v.createdAt);
    if (vDate.getTime() >= now - 7 * dayMs) {
      const hour = vDate.getHours();
      hourMap.set(hour, (hourMap.get(hour) || 0) + 1);
    }
    if (vDate.getTime() >= now - 30 * dayMs) {
      if (v.city) {
        const k = `${v.city}_${v.country}`;
        const prev = cityMap.get(k) || { city: v.city, country: v.country, count: 0 };
        prev.count += 1;
        cityMap.set(k, prev);
      }
      if (v.country) {
        countryMap.set(v.country, (countryMap.get(v.country) || 0) + 1);
      }
      const dayStr = vDate.toISOString().slice(0, 10);
      dayMap.set(dayStr, (dayMap.get(dayStr) || 0) + 1);
    }
  }

  res.json({
    total: visits.length,
    today: todayCount,
    last7Days: last7Count,
    last30Days: last30Count,
    hourlyDistribution: Array.from(hourMap.entries()).map(([hour, count]) => ({ hour, count })),
    topCities: Array.from(cityMap.values()).sort((a, b) => b.count - a.count).slice(0, 10),
    topCountries: Array.from(countryMap.entries())
      .map(([country, count]) => ({ country, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10),
    dailySeries: Array.from(dayMap.entries())
      .map(([day, count]) => ({ day, count }))
      .sort((a, b) => a.day.localeCompare(b.day)),
  });
});

apiRouter.get("/admin/sales/stats", (req, res) => {
  const period = req.query.period === "monthly" ? "monthly" : "weekly";
  const days = period === "monthly" ? 30 : 7;
  const since = Date.now() - days * 24 * 3600 * 1000;
  const orders = dbManager.getOrders();

  const windowOrders = orders.filter((o) => new Date(o.createdAt).getTime() >= since);
  const delivered = windowOrders.filter((o) => o.status === "Entregado");
  const cancelled = windowOrders.filter((o) => o.status === "Anulado");

  const totalRevenue = delivered.reduce((acc, o) => acc + (Number(o.total) || 0), 0);

  const productStats = new Map<string, { qty: number; revenue: number }>();
  const hourMap = new Map<number, { orders: number; revenue: number }>();
  const dowMap = new Map<number, { orders: number; revenue: number }>();
  const dayMap = new Map<string, { orders: number; revenue: number }>();

  for (let h = 0; h < 24; h++) hourMap.set(h, { orders: 0, revenue: 0 });
  for (let d = 0; d < 7; d++) dowMap.set(d, { orders: 0, revenue: 0 });

  for (const o of delivered) {
    const oDate = new Date(o.createdAt);
    const hour = oDate.getHours();
    const dow = oDate.getDay();
    const day = oDate.toISOString().slice(0, 10);
    const rev = Number(o.total) || 0;

    const hEntry = hourMap.get(hour)!;
    hEntry.orders += 1;
    hEntry.revenue += rev;

    const dEntry = dowMap.get(dow)!;
    dEntry.orders += 1;
    dEntry.revenue += rev;

    const dayEntry = dayMap.get(day) || { orders: 0, revenue: 0 };
    dayEntry.orders += 1;
    dayEntry.revenue += rev;
    dayMap.set(day, dayEntry);

    for (const it of o.items || []) {
      const q = Number(it.quantity) || 1;
      const p = Number(it.price) || 0;
      const prev = productStats.get(it.name) || { qty: 0, revenue: 0 };
      prev.qty += q;
      prev.revenue += q * p;
      productStats.set(it.name, prev);
    }
  }

  const topProducts = Array.from(productStats.entries())
    .map(([name, data]) => ({ name, qty: data.qty, revenue: data.revenue }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 10);

  res.json({
    period,
    days,
    totalOrders: delivered.length,
    cancelledOrders: cancelled.length,
    totalRevenue,
    topProducts,
    hourlyDistribution: Array.from(hourMap.entries()).map(([hour, val]) => ({
      hour,
      orders: val.orders,
      revenue: val.revenue,
    })),
    dayOfWeekDistribution: Array.from(dowMap.entries()).map(([dayOfWeek, val]) => ({
      dayOfWeek,
      orders: val.orders,
      revenue: val.revenue,
    })),
    dailySeries: Array.from(dayMap.entries())
      .map(([day, val]) => ({ day, orders: val.orders, revenue: val.revenue }))
      .sort((a, b) => a.day.localeCompare(b.day)),
  });
});

apiRouter.get("/admin/customers", (_req, res) => {
  const orders = dbManager.getOrders();
  const customerMap = new Map<string, { customer_name: string; phone: string; order_count: number; total_spent: number }>();

  for (const o of orders) {
    if (o.status === "Anulado") continue;
    const phone = o.phone || "Sin teléfono";
    const prev = customerMap.get(phone) || {
      customer_name: o.customerName || "Cliente",
      phone,
      order_count: 0,
      total_spent: 0,
    };
    prev.order_count += 1;
    prev.total_spent += Number(o.total) || 0;
    customerMap.set(phone, prev);
  }

  const topBuyers = Array.from(customerMap.values())
    .sort((a, b) => b.total_spent - a.total_spent)
    .slice(0, 20)
    .map((c) => ({
      name: c.customer_name,
      phone: c.phone,
      orderCount: c.order_count,
      totalSpent: c.total_spent,
    }));

  res.json({
    registered: [],
    topBuyers,
  });
});

// 10. Visits
apiRouter.post("/visits", async (req, res) => {
  const ip = (req.headers["x-forwarded-for"] as string)?.split(",")[0]?.trim() || req.socket.remoteAddress || "";
  const geo = await lookupIpGeo(ip);
  const pathStr = String(req.body.path || "/");
  const userAgent = String(req.headers["user-agent"] || "");

  const visit = dbManager.addVisit({
    ipAddress: ip,
    city: geo.city,
    region: geo.region,
    country: geo.country,
    path: pathStr,
    userAgent,
  });

  res.json({ ok: true, id: visit.id });
});

// 11. Delivery locations
apiRouter.get("/delivery-locations", (_req, res) => {
  res.json(dbManager.getDeliveryLocations().filter((l) => l.active));
});

apiRouter.get("/admin/delivery-locations", (_req, res) => {
  res.json(dbManager.getDeliveryLocations());
});

apiRouter.post("/admin/delivery-locations", (req, res) => {
  const body = req.body;
  const created = dbManager.addDeliveryLocation({
    name: body.name,
    price: Number(body.price) || 0,
    active: body.active !== undefined ? Boolean(body.active) : true,
    position: dbManager.getDeliveryLocations().length,
  });
  res.status(201).json(created);
});

apiRouter.patch("/admin/delivery-locations/:id", (req, res) => {
  const updated = dbManager.updateDeliveryLocation(Number(req.params.id), req.body);
  if (!updated) {
    res.status(404).json({ error: "Location not found" });
    return;
  }
  res.json(updated);
});

apiRouter.delete("/admin/delivery-locations/:id", (req, res) => {
  const ok = dbManager.deleteDeliveryLocation(Number(req.params.id));
  res.json({ ok });
});

// 12. Discounts
apiRouter.post("/discounts/apply", (req, res) => {
  const { code, orderTotal } = req.body;
  if (!code) {
    res.status(400).json({ error: "Código requerido" });
    return;
  }
  const discount = dbManager.getDiscountByCode(code);
  if (!discount) {
    res.status(404).json({ error: "Cupón no válido o inactivo" });
    return;
  }
  if (discount.minOrder && Number(orderTotal) < discount.minOrder) {
    res.status(400).json({
      error: `El pedido mínimo para este cupón es $${discount.minOrder.toLocaleString("es-CL")}`,
    });
    return;
  }
  let discountAmount = 0;
  if (discount.type === "percentage") {
    discountAmount = Math.round((Number(orderTotal) * discount.amount) / 100);
  } else {
    discountAmount = discount.amount;
  }
  res.json({
    valid: true,
    code: discount.code,
    type: discount.type,
    amount: discount.amount,
    discountAmount,
  });
});

apiRouter.get("/admin/discounts", (_req, res) => {
  res.json(dbManager.getDiscounts());
});

apiRouter.post("/admin/discounts", (req, res) => {
  const body = req.body;
  const created = dbManager.addDiscount({
    code: String(body.code).toUpperCase().trim(),
    type: body.type || "percentage",
    amount: Number(body.amount) || 0,
    customerId: body.customerId ?? null,
    active: body.active !== undefined ? Boolean(body.active) : true,
    minOrder: Number(body.minOrder) || 0,
    usesLeft: body.usesLeft !== undefined ? Number(body.usesLeft) : null,
    expiresAt: body.expiresAt ?? null,
  });
  res.status(201).json(created);
});

apiRouter.delete("/admin/discounts/:id", (req, res) => {
  const ok = dbManager.deleteDiscount(Number(req.params.id));
  res.json({ ok });
});

// 13. Customers
apiRouter.post("/customers/register", async (req, res) => {
  const { email, password, name, phone } = req.body;
  if (!email || !password) {
    res.status(400).json({ error: "Email y contraseña requeridos" });
    return;
  }
  const existing = dbManager.getCustomerByEmail(email);
  if (existing) {
    res.status(409).json({ error: "El email ya está registrado" });
    return;
  }
  const passwordHash = await bcrypt.hash(password, 10);
  const customer = dbManager.addCustomer({
    email,
    passwordHash,
    name: name || "",
    phone: phone || "",
  });
  const token = jwt.sign({ customerId: customer.id }, JWT_SECRET, { expiresIn: "30d" });
  res.status(201).json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone }, token });
});

apiRouter.post("/customers/login", async (req, res) => {
  const { email, password } = req.body;
  const customer = dbManager.getCustomerByEmail(email);
  if (!customer) {
    res.status(401).json({ error: "Credenciales inválidas" });
    return;
  }
  const valid = await bcrypt.compare(password, customer.passwordHash);
  if (!valid) {
    res.status(401).json({ error: "Credenciales inválidas" });
    return;
  }
  const token = jwt.sign({ customerId: customer.id }, JWT_SECRET, { expiresIn: "30d" });
  res.json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone }, token });
});

apiRouter.get("/customers/me", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const customer = dbManager.getCustomerById(customerId);
  if (!customer) {
    res.status(404).json({ error: "Cliente no encontrado" });
    return;
  }
  res.json({ customer: { id: customer.id, email: customer.email, name: customer.name, phone: customer.phone } });
});

apiRouter.get("/customers/cart", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const cart = dbManager.getSavedCart(customerId);
  res.json({ cart: cart ? cart.items : [] });
});

apiRouter.post("/customers/cart", (req, res) => {
  const customerId = getCustomerIdFromReq(req);
  if (!customerId) {
    res.status(401).json({ error: "No autenticado" });
    return;
  }
  const items = req.body.items || [];
  const saved = dbManager.saveCart(customerId, "Mi Carrito", items);
  res.json({ ok: true, cart: saved.items });
});

// 14. File Storage for images
apiRouter.post("/storage/uploads/request-url", (_req, res) => {
  const fileId = randomUUID();
  res.json({
    uploadURL: `/api/storage/uploads/${fileId}`,
    objectPath: `/objects/${fileId}`,
  });
});

apiRouter.put("/storage/uploads/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  const writeStream = fs.createWriteStream(filePath);
  req.pipe(writeStream);
  writeStream.on("finish", () => {
    res.status(200).json({ ok: true });
  });
  writeStream.on("error", (err) => {
    console.error("Upload error:", err);
    res.status(500).json({ error: "Upload failed" });
  });
});

apiRouter.get("/storage/objects/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  if (fs.existsSync(filePath)) {
    res.sendFile(filePath);
  } else {
    res.status(404).json({ error: "Object not found" });
  }
});

apiRouter.delete("/storage/objects/:id", (req, res) => {
  const fileId = req.params.id;
  const filePath = path.join(UPLOADS_DIR, fileId);
  if (fs.existsSync(filePath)) {
    try {
      fs.unlinkSync(filePath);
    } catch {}
  }
  res.json({ ok: true });
});

// 15. Resolve OpenGraph image
apiRouter.get("/resolve-image", async (req, res) => {
  const url = req.query.url as string;
  if (!url) {
    res.status(400).json({ error: "url param required" });
    return;
  }
  try {
    const response = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0" },
      signal: AbortSignal.timeout(8000),
    });
    const html = await response.text();
    const ogMatch =
      html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i) ||
      html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);
    if (ogMatch?.[1]) {
      res.json({ resolved: ogMatch[1] });
      return;
    }
    res.status(422).json({ error: "No direct image found" });
  } catch {
    res.status(502).json({ error: "Failed to fetch URL" });
  }
});

// 16. AI Copywriting with Gemini
apiRouter.post("/ai/improve-banner", async (req, res) => {
  const { title = "", subtitle = "" } = req.body;
  if (!title.trim() && !subtitle.trim()) {
    res.status(400).json({ error: "Empty input" });
    return;
  }

  const prompt = `Eres un copywriter experto en marketing juvenil para una botillería chilena llamada "Fella's Market" en Alerce, Puerto Montt.
Mejora el siguiente título y subtítulo de un banner promocional:
Título actual: "${title}"
Subtítulo actual: "${subtitle}"

Reglas:
- Tono juvenil, llamativo, fiestero y con buena onda chilena.
- TÍTULO: máximo 4-6 palabras, impactante, en MAYÚSCULAS con 1-2 emojis.
- SUBTÍTULO: máximo 10-14 palabras, frase atractiva con hashtag corto (#Alerce, #FindeSemana, #Promos).
- Responde estrictamente en formato JSON: {"title": "...", "subtitle": "..."}`;

  const apiKey = process.env.GEMINI_API_KEY;
  if (apiKey) {
    try {
      const ai = new GoogleGenAI({ apiKey });
      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: prompt,
        config: { responseMimeType: "application/json" },
      });
      const parsed = JSON.parse(response.text || "{}");
      if (parsed.title && parsed.subtitle) {
        res.json({ title: parsed.title, subtitle: parsed.subtitle });
        return;
      }
    } catch (err) {
      console.warn("Gemini call fallback:", err);
    }
  }

  // High quality fallback
  res.json({
    title: (title || "FELLA'S MARKET EN VIVO 🍻").toUpperCase(),
    subtitle: (subtitle || "Las mejores promos y delivery a la puerta de tu casa #Alerce"),
  });
});
