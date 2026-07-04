import bcrypt from "bcryptjs";
import cors from "cors";
import { randomUUID } from "crypto";
import dotenv from "dotenv";
import express from "express";
import fs from "fs";
import jwt from "jsonwebtoken";
import mongoose from "mongoose";
import multer from "multer";
import path from "path";
import XLSX from "xlsx";

dotenv.config();

const app = express();
const port = process.env.PORT || 5000;
const jwtSecret = process.env.JWT_SECRET || "development-secret";

const ROLE_ASSOCIATE_MEMBER = "associate-member";
const ROLE_ADMIN = "admin";
const ROLE_SUPER_ADMIN = "super-admin";

const STATUS_PENDING = "pending-review";
const STATUS_ACTIVE = "active";
const STATUS_SUSPENDED = "suspended";

const ADMIN_STATUS_ACTIVE = "Active";
const ADMIN_STATUS_DEACTIVE = "Deactive";

const ORDER_STATUS_PENDING = "pending";
const ORDER_STATUS_PRINTING = "printing";
const ORDER_STATUS_PACKAGING = "packaging";
const ORDER_STATUS_DISPATCHED = "dispatched";
const ORDER_STATUS_CANCELLED = "cancelled";
const ORDER_STATUS_IMPROPER = "improper";
const ORDER_STATUS_COMPLETED = "completed";
const ORDER_STATUS_REJECTED = "rejected";
const PENDING_ORDER_ROUTE_STATUSES = [ORDER_STATUS_PENDING];
const PRINTING_ORDER_ROUTE_STATUSES = [ORDER_STATUS_PRINTING];
const PACKAGING_ORDER_ROUTE_STATUSES = [ORDER_STATUS_PACKAGING];
const DISPATCHED_ORDER_ROUTE_STATUSES = [ORDER_STATUS_DISPATCHED];
const COMPLETED_ORDER_ROUTE_STATUSES = [ORDER_STATUS_COMPLETED];
const CANCELLED_ORDER_ROUTE_STATUSES = [ORDER_STATUS_CANCELLED];
const IMPROPER_ORDER_ROUTE_STATUSES = [ORDER_STATUS_IMPROPER];
const REJECTED_ORDER_ROUTE_STATUSES = [ORDER_STATUS_REJECTED];

const DESIGN_SOURCE_ONLINE = "online-upload";
const DESIGN_SOURCE_EMAIL = "email";

const superAdminCredentials = {
  country: "India",
  mobileNumber: "9975813249",
  password: "SuperAdmin@123",
  role: ROLE_SUPER_ADMIN,
  status: STATUS_ACTIVE,
  name: "Super Admin"
};

app.use(cors());
app.use(express.json({ limit: "10mb" }));

const uploadsRootPath = path.join(process.cwd(), "uploads");
const designUploadsPath = path.join(uploadsRootPath, "design");
fs.mkdirSync(designUploadsPath, { recursive: true });
app.use("/uploads", express.static(uploadsRootPath));

const designUpload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, callback) => callback(null, designUploadsPath),
    filename: (_req, file, callback) => {
      const ext = path.extname(file.originalname || "");
      callback(null, `${Date.now()}-${randomUUID()}${ext}`);
    }
  }),
  limits: { fileSize: 100 * 1024 * 1024 },
  fileFilter: (_req, file, callback) => {
    const extension = path.extname(file.originalname || "").slice(1).toLowerCase();
    const allowed = new Set(["pdf", "cdr", "ai", "psd", "jpeg", "jpg", "png"]);
    if (!allowed.has(extension)) {
      callback(new Error("Invalid file format. Please upload AI, PDF, CDR, PSD, JPEG, JPG, or PNG."));
      return;
    }
    callback(null, true);
  }
});

const userSchema = new mongoose.Schema(
  {
    businessName: { type: String, required: true, trim: true },
    ownerName: { type: String, required: true, trim: true },
    mobile: { type: String, required: true, trim: true },
    email: { type: String, required: true, trim: true, lowercase: true },
    passwordHash: { type: String, required: true },
    referenceCode: { type: String, trim: true, default: "" },
    country: { type: String, required: true, trim: true },
    state: { type: String, required: true, trim: true },
    district: { type: String, required: true, trim: true },
    city: { type: String, required: true, trim: true },
    pinCode: { type: String, required: true, trim: true },
    gstNumber: { type: String, trim: true, default: "" },
    address: { type: String, required: true, trim: true },
    services: [{ type: String }],
    termsAccepted: { type: Boolean, default: false },
    captchaVerified: { type: Boolean, default: false },
    role: { type: String, default: ROLE_ASSOCIATE_MEMBER, index: true },
    status: { type: String, default: STATUS_PENDING, index: true },
    associateMemberAccessEnabled: { type: Boolean, default: false, index: true },
    assignedAdminId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null, index: true },
    lastLoginAt: { type: Date, default: null },
    lastActivityAt: { type: Date, default: null }
  },
  { timestamps: true }
);

userSchema.index({ country: 1, mobile: 1 }, { unique: true });
userSchema.index({ email: 1 }, { unique: true });

const orderSchema = new mongoose.Schema(
  {
    orderNumber: { type: Number, default: null, index: true },
    orderName: { type: String, trim: true, default: "" },
    customerName: { type: String, trim: true, default: "" },
    customerMobile: { type: String, trim: true, default: "" },
    orderDetailsOverview: { type: String, trim: true, default: "" },
    bagName: { type: String, trim: true, default: "" },
    printSide: { type: String, trim: true, default: "" },
    quantity: { type: Number, default: 0 },
    bagSize: { type: String, trim: true, default: "" },
    bagColor: { type: String, trim: true, default: "" },
    textColorType: { type: String, trim: true, default: "" },
    textColors: [{ type: String, trim: true }],
    printingPress: { type: String, trim: true, default: "" },
    privacy: { type: String, trim: true, default: "" },
    deliveryOption: { type: String, trim: true, default: "" },
    fileOption: { type: String, trim: true, default: "" },
    sellingPrice: { type: Number, default: 0 },
    remark: { type: String, trim: true, default: "" },
    pressline: { type: String, trim: true, default: "" },
    orderedAt: { type: Date, default: Date.now, index: true },
    status: { type: String, default: ORDER_STATUS_PENDING, index: true },
    designSubmissionSource: { type: String, default: DESIGN_SOURCE_ONLINE, index: true },
    designFileName: { type: String, trim: true, default: "" },
    designFileType: { type: String, trim: true, default: "" },
    designFileUrl: { type: String, trim: true, default: "" },
    courierName: { type: String, trim: true, default: "" },
    courierTrackingNumber: { type: String, trim: true, default: "" },
    courierTrackingUrl: { type: String, trim: true, default: "" },
    dispatchDateTime: { type: Date, default: null, index: true },
    dispatchNotes: { type: String, trim: true, default: "" },
    deliveryStatus: { type: String, trim: true, default: "" },
    deliveryDateTime: { type: Date, default: null },
    isUrgent: { type: Boolean, default: false, index: true },
    placedByUserId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null },
    assignedAssociateMemberId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null, index: true },
    assignedAdminId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null, index: true },
    referenceNo: { type: String, trim: true, default: "" },
    basePayableAmount: { type: Number, default: 0 },
    pdfDiscountAmount: { type: Number, default: 0 },
    walletDebitAmount: { type: Number, default: 0 },
    statusHistory: [
      {
        status: { type: String, default: ORDER_STATUS_PENDING },
        note: { type: String, trim: true, default: "" },
        changedAt: { type: Date, default: Date.now },
        changedByRole: { type: String, default: "" },
        changedById: { type: String, default: "" },
        changedByName: { type: String, default: "" }
      }
    ],
    dispatchHistory: [
      {
        status: { type: String, default: ORDER_STATUS_DISPATCHED },
        note: { type: String, trim: true, default: "" },
        courierName: { type: String, trim: true, default: "" },
        courierTrackingNumber: { type: String, trim: true, default: "" },
        courierTrackingUrl: { type: String, trim: true, default: "" },
        deliveryStatus: { type: String, trim: true, default: "" },
        eventAt: { type: Date, default: Date.now },
        updatedByRole: { type: String, default: "" },
        updatedById: { type: String, default: "" },
        updatedByName: { type: String, default: "" }
      }
    ]
  },
  { timestamps: true }
);

const topUpRequestSchema = new mongoose.Schema(
  {
    status: { type: String, default: "pending", index: true },
    requestedByUserId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null },
    requestedByExternalKey: { type: String, trim: true, default: "", index: true },
    requestedByRole: { type: String, trim: true, default: "" },
    requestedByDisplayName: { type: String, trim: true, default: "" },
    amount: { type: Number, default: 0 },
    remarks: { type: String, trim: true, default: "" },
    reviewedByRole: { type: String, trim: true, default: "" },
    reviewedById: { type: String, trim: true, default: "" },
    reviewedByName: { type: String, trim: true, default: "" },
    reviewedAt: { type: Date, default: null },
    decisionNote: { type: String, trim: true, default: "" }
  },
  { timestamps: true }
);

const auditLogSchema = new mongoose.Schema(
  {
    actorRole: { type: String, required: true },
    actorId: { type: String, required: true },
    actorName: { type: String, required: true },
    module: { type: String, required: true },
    action: { type: String, required: true },
    targetType: { type: String, required: true },
    targetId: { type: String, default: "" },
    targetName: { type: String, default: "" },
    details: { type: mongoose.Schema.Types.Mixed, default: {} }
  },
  { timestamps: true }
);

const User = mongoose.model("User", userSchema);
const Order = mongoose.model("Order", orderSchema);
const TopUpRequest = mongoose.model("TopUpRequest", topUpRequestSchema);
const AuditLog = mongoose.model("AuditLog", auditLogSchema);

function createToken(payload) {
  return jwt.sign(payload, jwtSecret, { expiresIn: "7d" });
}

function getBearerToken(authorizationHeader) {
  if (!authorizationHeader) return null;
  const [scheme, token] = authorizationHeader.split(" ");
  if (scheme !== "Bearer" || !token) return null;
  return token;
}

function ensureDatabaseConnected(_req, res, next) {
  if (mongoose.connection.readyState !== 1) {
    return res.status(503).json({ message: "Database is not connected." });
  }
  return next();
}

function authenticate(req, res, next) {
  const token = getBearerToken(req.headers.authorization);

  if (!token) {
    return res.status(401).json({ message: "Authorization token is required." });
  }

  try {
    req.auth = jwt.verify(token, jwtSecret);

    if (req.auth?.sub && req.auth.sub !== "super-admin" && mongoose.connection.readyState === 1) {
      User.updateOne({ _id: req.auth.sub }, { $set: { lastActivityAt: new Date() } }).catch(() => {});
    }

    return next();
  } catch (_error) {
    return res.status(401).json({ message: "Invalid or expired token." });
  }
}

function requireRole(role) {
  return (req, res, next) => {
    if (req.auth?.role !== role) {
      return res.status(403).json({ message: "You do not have permission to access this resource." });
    }

    return next();
  };
}

app.post("/api/uploads/design", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), (req, res) => {
  designUpload.single("file")(req, res, (error) => {
    if (error) {
      return res.status(400).json({ message: error.message || "Failed to upload design file." });
    }

    if (!req.file) {
      return res.status(400).json({ message: "Design file is required." });
    }

    const fileUrl = `${req.protocol}://${req.get("host")}/uploads/design/${req.file.filename}`;
    const originalName = req.file.originalname || req.file.filename;
    const extension = path.extname(originalName).slice(1).toUpperCase();

    return res.json({
      message: "Design file uploaded successfully.",
      fileName: originalName,
      fileType: extension,
      fileUrl
    });
  });
});

app.post("/api/orders/:orderId/design", authenticate, ensureDatabaseConnected, (req, res) => {
  designUpload.single("file")(req, res, async (error) => {
    if (error) {
      return res.status(400).json({ message: error.message || "Failed to upload design file." });
    }

    if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
      return res.status(400).json({ message: "Invalid Order record ID." });
    }

    if (!req.file) {
      return res.status(400).json({ message: "Design file is required." });
    }

    const order = await Order.findById(req.params.orderId).select("_id placedByUserId assignedAdminId status orderNumber orderName");
    if (!order) {
      return res.status(404).json({ message: "Order record not found." });
    }

    const role = req.auth?.role;
    const actorId = String(req.auth?.sub || "");
    const placedById = order.placedByUserId ? order.placedByUserId.toString() : "";
    const assignedAdminId = order.assignedAdminId ? order.assignedAdminId.toString() : "";

    const canUpdate =
      role === ROLE_SUPER_ADMIN ||
      (role === ROLE_ADMIN && assignedAdminId && assignedAdminId === actorId) ||
      (role === ROLE_ASSOCIATE_MEMBER && placedById && placedById === actorId);

    if (!canUpdate) {
      return res.status(403).json({ message: "You do not have permission to update this order." });
    }

    const fileUrl = `${req.protocol}://${req.get("host")}/uploads/design/${req.file.filename}`;
    const originalName = req.file.originalname || req.file.filename;
    const extension = path.extname(originalName).slice(1).toUpperCase();
    const actorName = await getActorName(req.auth);
    const note = String(req.body.note || "").trim();
    const historyStatus = order.status || ORDER_STATUS_PENDING;

    const updatedOrder = await Order.findByIdAndUpdate(
      req.params.orderId,
      {
        $set: {
          designSubmissionSource: DESIGN_SOURCE_ONLINE,
          designFileName: originalName,
          designFileType: extension,
          designFileUrl: fileUrl,
          updatedAt: new Date()
        },
        $push: {
          statusHistory: {
            status: historyStatus,
            note: note || "Design file uploaded/updated.",
            changedAt: new Date(),
            changedByRole: req.auth.role,
            changedById: req.auth.sub,
            changedByName: actorName
          }
        }
      },
      { new: true }
    );

    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Order Design Upload",
      action: "Upload/Replace Design File",
      targetType: "Order",
      targetId: updatedOrder._id.toString(),
      targetName: updatedOrder.orderName || `Order ${updatedOrder.orderNumber}`,
      details: {
        orderNumber: updatedOrder.orderNumber,
        fileName: originalName,
        fileUrl
      }
    });

    return res.json({
      message: "Design file attached to order successfully.",
      order: await buildOrderDetails(updatedOrder)
    });
  });
});

function getModulePath(role) {
  if (role === ROLE_SUPER_ADMIN) return "/dashboard/super-admin";
  if (role === ROLE_ADMIN) return "/dashboard/admin";
  return "/dashboard/associate-member";
}

function buildUserResponse(user) {
  const walletBalance = Number(user.walletBalance);
  return {
    id: user._id?.toString?.() || "super-admin",
    country: user.country,
    mobileNumber: user.mobile || user.mobileNumber,
    role: user.role,
    status: user.status,
    ownerName: user.ownerName || user.name,
    businessName: user.businessName || "Printers Club Admin",
    modulePath: getModulePath(user.role),
    walletBalance: Number.isFinite(walletBalance) ? walletBalance.toFixed(2) : "0.00"
  };
}

async function getAssociateWalletBalance(memberId) {
  if (!memberId) return 0;

  const normalizedMemberId =
    typeof memberId === "string" && mongoose.isValidObjectId(memberId)
      ? new mongoose.Types.ObjectId(memberId)
      : memberId;

  const [approvedTopUps, orderDebits] = await Promise.all([
    TopUpRequest.aggregate([
      { $match: { requestedByUserId: normalizedMemberId, status: "approved" } },
      { $group: { _id: null, total: { $sum: "$amount" } } }
    ]),
    Order.aggregate([
      { $match: { placedByUserId: normalizedMemberId } },
      { $group: { _id: null, total: { $sum: "$walletDebitAmount" } } }
    ])
  ]);

  const topUpTotal = Number(approvedTopUps[0]?.total || 0);
  const debitTotal = Number(orderDebits[0]?.total || 0);
  return Math.max(0, Number((topUpTotal - debitTotal).toFixed(2)));
}

async function getWalletBalanceByExternalKey(externalKey) {
  const normalizedKey = String(externalKey || "").trim();
  if (!normalizedKey) return 0;

  const approvedTopUps = await TopUpRequest.aggregate([
    { $match: { requestedByExternalKey: normalizedKey, status: "approved" } },
    { $group: { _id: null, total: { $sum: "$amount" } } }
  ]);

  return Math.max(0, Number(Number(approvedTopUps[0]?.total || 0).toFixed(2)));
}

async function getWalletBalanceForActor({ userId = "", role = "", externalKey = "" } = {}) {
  if (String(role || "").trim() === ROLE_SUPER_ADMIN) {
    return getWalletBalanceByExternalKey(externalKey || "super-admin");
  }

  return getAssociateWalletBalance(userId);
}

function normalizeTopUpStatusLabel(statusValue) {
  const safeStatus = String(statusValue || "").trim();
  return safeStatus ? safeStatus.replace(/^\w/, (char) => char.toUpperCase()) : "--";
}

async function getWalletHistoryRequestsForActor({ userId = "", role = "", externalKey = "", limit = 100 } = {}) {
  const cappedLimit = Math.min(Math.max(Number(limit) || 100, 1), 100);

  if (String(role || "").trim() === ROLE_SUPER_ADMIN) {
    const normalizedKey = String(externalKey || "super-admin").trim();
    if (!normalizedKey) return [];

    return TopUpRequest.find({ requestedByExternalKey: normalizedKey }).sort({ createdAt: -1 }).limit(cappedLimit);
  }

  if (!userId) return [];

  const normalizedUserId =
    typeof userId === "string" && mongoose.isValidObjectId(userId) ? new mongoose.Types.ObjectId(userId) : userId;

  return TopUpRequest.find({ requestedByUserId: normalizedUserId }).sort({ createdAt: -1 }).limit(cappedLimit);
}

function buildWalletHistoryResponse(requests = []) {
  return {
    meta: {
      columns: [
        { key: "reference", label: "Reference" },
        { key: "amount", label: "Amount" },
        { key: "status", label: "Status" },
        { key: "requestedOn", label: "Requested On" }
      ]
    },
    summary: {
      total: requests.length,
      pending: requests.filter((request) => request.status === "pending").length,
      approved: requests.filter((request) => request.status === "approved").length,
      rejected: requests.filter((request) => request.status === "rejected").length
    },
    items: requests.map((request) => ({
      id: request._id.toString(),
      reference: `TPR-${request._id.toString().slice(-6).toUpperCase()}`,
      amount: Number(request.amount || 0).toFixed(2),
      status: normalizeTopUpStatusLabel(request.status),
      requestedOn: formatDateTime(request.createdAt)
    }))
  };
}

function normalizeRoleLabel(roleValue) {
  const safe = String(roleValue || "").trim();
  if (!safe) return "--";
  return safe
    .split("-")
    .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
    .join(" ");
}

async function buildWalletRequestDetails(requestDocument) {
  if (!requestDocument) return null;

  const requestId = requestDocument._id?.toString?.() || "";
  const requester =
    requestDocument.requestedByUserId && typeof requestDocument.requestedByUserId === "object" && requestDocument.requestedByUserId._id
      ? requestDocument.requestedByUserId
      : requestDocument.requestedByUserId
        ? await User.findById(requestDocument.requestedByUserId).select(
            "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt"
          )
        : null;

  const requesterName =
    requester?.ownerName ||
    String(requestDocument.requestedByDisplayName || "").trim() ||
    (String(requestDocument.requestedByExternalKey || "").trim() === "super-admin" ? "Super Admin" : "--");

  return {
    id: requestId,
    reference: `TPR-${requestId.slice(-6).toUpperCase()}`,
    amount: Number(requestDocument.amount || 0).toFixed(2),
    status: normalizeTopUpStatusLabel(requestDocument.status),
    statusValue: String(requestDocument.status || "").trim().toLowerCase(),
    remarks: String(requestDocument.remarks || "").trim(),
    requestedOn: formatDateTime(requestDocument.createdAt),
    updatedOn: formatDateTime(requestDocument.updatedAt),
    reviewedOn: formatDateTime(requestDocument.reviewedAt),
    reviewedByName: String(requestDocument.reviewedByName || "").trim(),
    reviewedByRole: normalizeRoleLabel(requestDocument.reviewedByRole),
    decisionNote: String(requestDocument.decisionNote || "").trim(),
    requestedByDisplayName: requesterName,
    requestedByRole: normalizeRoleLabel(requestDocument.requestedByRole || requester?.role || ""),
    requestedByExternalKey: String(requestDocument.requestedByExternalKey || "").trim(),
    requester: requester
      ? {
          id: requester._id?.toString?.() || "",
          name: requester.ownerName || "--",
          mobileNumber: requester.mobile || "--",
          email: requester.email || "--",
          businessName: requester.businessName || "--",
          address: requester.address || "--",
          country: requester.country || "--",
          state: requester.state || "--",
          district: requester.district || "--",
          city: requester.city || "--",
          pinCode: requester.pinCode || "--",
          gstNumber: requester.gstNumber || "--",
          role: normalizeRoleLabel(requester.role || ""),
          status: normalizeAdminStatusLabel(requester.status),
          registrationDate: formatDateTime(requester.createdAt)
        }
      : null
  };
}

function buildWalletManagementListRow(details = {}) {
  return {
    id: details.id,
    reference: details.reference,
    module: "Wallet Management",
    userOrder: details.requester?.name || details.requestedByDisplayName || "--",
    requestedByRole: details.requestedByRole,
    businessName: details.requester?.businessName || "--",
    amount: details.amount,
    status: String(details.statusValue || "").trim().toLowerCase() || "pending",
    statusLabel: details.status,
    updatedBy: details.reviewedByName || details.requestedByDisplayName || "--",
    updatedOn: details.reviewedOn !== "-" ? details.reviewedOn : details.updatedOn,
    requestedOn: details.requestedOn
  };
}

async function listWalletRequestsForSuperAdmin(view = "transactions") {
  const normalizedView = String(view || "transactions").trim().toLowerCase();
  const query = {};

  if (normalizedView === "pending") {
    query.status = "pending";
  } else if (normalizedView === "approved") {
    query.status = "approved";
  } else if (normalizedView === "rejected") {
    query.status = "rejected";
  }

  const requests = await TopUpRequest.find(query)
    .sort({ createdAt: -1 })
    .limit(100)
    .populate("requestedByUserId", "ownerName mobile businessName");

  return {
    meta: {
      columns: [
        { key: "reference", label: "Reference" },
        { key: "memberName", label: "Member Name" },
        { key: "mobile", label: "Mobile" },
        { key: "amount", label: "Amount" },
        { key: "status", label: "Status" },
        { key: "businessName", label: "Business / Firm" }
      ]
    },
    summary: {
      total: requests.length,
      pending: requests.filter((request) => request.status === "pending").length,
      approved: requests.filter((request) => request.status === "approved").length,
      rejected: requests.filter((request) => request.status === "rejected").length
    },
    items: requests.map((request) => ({
      id: request._id.toString(),
      reference: `TPR-${request._id.toString().slice(-6).toUpperCase()}`,
      memberName: request.requestedByUserId?.ownerName || request.requestedByDisplayName || "--",
      mobile: request.requestedByUserId?.mobile || "--",
      amount: Number(request.amount || 0).toFixed(2),
      status: normalizeTopUpStatusLabel(request.status),
      businessName: request.requestedByUserId?.businessName || "--"
    }))
  };
}

async function listWalletTransactionsForSuperAdmin() {
  const [requests, orders] = await Promise.all([
    TopUpRequest.find({})
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(200)
      .populate("requestedByUserId", "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt"),
    Order.find({ walletDebitAmount: { $gt: 0 } })
      .sort({ orderedAt: -1, createdAt: -1 })
      .limit(200)
      .populate("placedByUserId", "ownerName mobile email businessName role status")
      .populate("assignedAdminId", "ownerName mobile email businessName role status")
      .select("orderNumber orderName status orderedAt referenceNo walletDebitAmount placedByUserId assignedAdminId createdAt updatedAt")
  ]);

  const creditItems = [];
  for (const request of requests) {
    const details = await buildWalletRequestDetails(request);
    creditItems.push({
      id: `wallet-credit-${details.id}`,
      sourceId: details.id,
      reference: details.reference,
      transactionType: "credit",
      transactionTypeLabel: "Credit",
      sourceModule: "Wallet Top-Up",
      actorName: details.requestedByDisplayName || "--",
      actorRole: details.requestedByRole || "--",
      businessName: details.requester?.businessName || "--",
      mobileNumber: details.requester?.mobileNumber || "--",
      creditAmount: Number(details.amount || 0).toFixed(2),
      debitAmount: "0.00",
      amount: Number(details.amount || 0).toFixed(2),
      status: details.status,
      statusValue: details.statusValue || "pending",
      transactionDate: details.requestedOn,
      updatedOn: details.updatedOn,
      reviewedByName: details.reviewedByName || "--",
      reviewedByRole: details.reviewedByRole || "--",
      decisionNote: details.decisionNote || "--",
      remarks: details.remarks || "--",
      flowLabel: `${details.requestedByRole || "User"} -> Wallet`,
      detailsPath: `/dashboard/super-admin/wallet-management/details/${details.id}`,
      dateValue: new Date(request.createdAt || request.updatedAt || Date.now()).toISOString(),
      updatedValue: new Date(request.updatedAt || request.createdAt || Date.now()).toISOString()
    });
  }

  const debitItems = orders.map((order) => {
    const orderId = order._id?.toString?.() || "";
    const placedByUser =
      order.placedByUserId && typeof order.placedByUserId === "object" && order.placedByUserId._id ? order.placedByUserId : null;
    const assignedAdmin =
      order.assignedAdminId && typeof order.assignedAdminId === "object" && order.assignedAdminId._id ? order.assignedAdminId : null;
    const actorRole = normalizeRoleLabel(placedByUser?.role || "");
    const debitAmount = Number(order.walletDebitAmount || 0).toFixed(2);

    return {
      id: `wallet-debit-${orderId}`,
      sourceId: orderId,
      reference: order.referenceNo || `ORD-${String(order.orderNumber || orderId.slice(-6)).toUpperCase()}`,
      transactionType: "debit",
      transactionTypeLabel: "Debit",
      sourceModule: "Order Management",
      actorName: placedByUser?.ownerName || "--",
      actorRole: actorRole || "--",
      businessName: placedByUser?.businessName || "--",
      mobileNumber: placedByUser?.mobile || "--",
      creditAmount: "0.00",
      debitAmount,
      amount: debitAmount,
      status: normalizeOrderStatusLabel(order.status),
      statusValue: String(order.status || "").trim().toLowerCase(),
      transactionDate: formatDateTime(order.orderedAt || order.createdAt),
      updatedOn: formatDateTime(order.updatedAt || order.createdAt),
      reviewedByName: assignedAdmin?.ownerName || "--",
      reviewedByRole: normalizeRoleLabel(assignedAdmin?.role || "") || "--",
      decisionNote: order.orderName || "--",
      remarks: order.orderName || "--",
      flowLabel: `${actorRole || "User"} -> Order Wallet Debit`,
      detailsPath: orderId ? `/dashboard/super-admin/order-management/details/${orderId}` : "",
      dateValue: new Date(order.orderedAt || order.createdAt || Date.now()).toISOString(),
      updatedValue: new Date(order.updatedAt || order.createdAt || Date.now()).toISOString()
    };
  });

  const items = [...creditItems, ...debitItems].sort((left, right) => {
    const leftDate = new Date(left.dateValue || left.updatedValue || 0).getTime();
    const rightDate = new Date(right.dateValue || right.updatedValue || 0).getTime();
    return rightDate - leftDate;
  });

  return {
    summary: {
      total: items.length,
      credits: creditItems.length,
      debits: debitItems.length,
      approved: creditItems.filter((item) => item.statusValue === "approved").length,
      rejected: creditItems.filter((item) => item.statusValue === "rejected").length
    },
    items
  };
}

async function buildWalletHistoryForActor(options = {}) {
  const requests = await getWalletHistoryRequestsForActor(options);
  return buildWalletHistoryResponse(requests);
}

function buildTopUpActor(auth) {
  const role = String(auth?.role || "").trim();
  if (role === ROLE_SUPER_ADMIN) {
    return {
      requestedByUserId: null,
      requestedByExternalKey: "super-admin",
      requestedByRole: ROLE_SUPER_ADMIN,
      requestedByDisplayName: "Super Admin"
    };
  }

  return {
    requestedByUserId: auth?.sub || null,
    requestedByExternalKey: "",
    requestedByRole: role,
    requestedByDisplayName: ""
  };
}

function normalizeAdminStatusLabel(statusValue) {
  return statusValue === STATUS_ACTIVE ? ADMIN_STATUS_ACTIVE : ADMIN_STATUS_DEACTIVE;
}

function getAdminDisplayId(mobileNumber) {
  const mobile = String(mobileNumber || "").trim();
  const digits = mobile.replace(/\D/g, "");
  return `ADM-${digits.slice(-4).padStart(4, "0")}`;
}

function buildAddress(user) {
  return [user.address, user.city, user.district, user.state, user.pinCode].filter(Boolean).join(", ");
}

function formatDateTime(value) {
  if (!value) return "--";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";

  return date.toLocaleString("en-IN", {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

const allowedRegistrationServices = new Set(["printingServices", "expo", "magazine", "advertiser"]);
const captchaChallenges = new Map();
const captchaExpiryMs = 5 * 60 * 1000;
const registrationEmailPattern = /^\S+@\S+\.\S+$/;
const registrationBusinessPattern = /^[a-zA-Z0-9.,\-/()\s]*$/;
const registrationPinPattern = /^[1-9][0-9]{5}$/;
const registrationGstPattern = /^([0][1-9]|[1-2][0-9]|[3][0-7])([A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])$/;

function generateCaptchaText(length = 6) {
  const characters = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  return Array.from({ length }, () => characters[Math.floor(Math.random() * characters.length)]).join("");
}

function pruneExpiredCaptchaChallenges() {
  const now = Date.now();
  for (const [challengeId, challenge] of captchaChallenges.entries()) {
    if (challenge.expiresAt <= now) {
      captchaChallenges.delete(challengeId);
    }
  }
}

function createCaptchaChallenge() {
  pruneExpiredCaptchaChallenges();
  const challengeId = randomUUID();
  const captchaText = generateCaptchaText();
  captchaChallenges.set(challengeId, {
    answer: captchaText,
    expiresAt: Date.now() + captchaExpiryMs
  });

  return {
    challengeId,
    captchaText,
    expiresInSeconds: Math.floor(captchaExpiryMs / 1000)
  };
}

function verifyCaptchaChallenge(challengeId, captchaValue) {
  pruneExpiredCaptchaChallenges();
  const challenge = captchaChallenges.get(String(challengeId || "").trim());

  if (!challenge) {
    return {
      ok: false,
      message: "CAPTCHA has expired. Please refresh and try again."
    };
  }

  captchaChallenges.delete(String(challengeId || "").trim());

  if (challenge.expiresAt <= Date.now()) {
    return {
      ok: false,
      message: "CAPTCHA has expired. Please refresh and try again."
    };
  }

  if (String(captchaValue || "").trim().toUpperCase() !== challenge.answer) {
    return {
      ok: false,
      message: "Incorrect CAPTCHA. Please enter the displayed text."
    };
  }

  return {
    ok: true
  };
}

function normalizeMobileNumber(country, mobileNumber) {
  const normalizedCountry = String(country || "").trim().toLowerCase();
  let digits = String(mobileNumber || "").replace(/\D/g, "");

  if (normalizedCountry === "india") {
    if (digits.startsWith("91") && digits.length === 12) {
      digits = digits.slice(2);
    }

    if (digits.startsWith("0") && digits.length === 11) {
      digits = digits.slice(1);
    }
  }

  return digits;
}

function validateRegistrationPayload(payload = {}) {
  const errors = [];
  const normalizedCountry = String(payload.country || "").trim();
  const normalizedMobile = normalizeMobileNumber(normalizedCountry, payload.mobile);
  const normalizedEmail = String(payload.email || "").trim().toLowerCase();
  const normalizedBusinessName = String(payload.businessName || "").trim();
  const normalizedOwnerName = String(payload.ownerName || "").trim();
  const normalizedState = String(payload.state || "").trim();
  const normalizedDistrict = String(payload.district || "").trim();
  const normalizedCity = String(payload.city || "").trim();
  const normalizedPinCode = String(payload.pinCode || "").trim();
  const normalizedAddress = String(payload.address || "").trim();
  const normalizedGstNumber = String(payload.gstNumber || "").trim().toUpperCase();
  const normalizedReferenceCode = String(payload.referenceCode || "").trim();
  const normalizedServices = Array.isArray(payload.services)
    ? [...new Set(payload.services.map((service) => String(service || "").trim()).filter(Boolean))]
    : [];

  if (!normalizedBusinessName) errors.push("Business / Firm Name is required.");
  if (normalizedBusinessName && !registrationBusinessPattern.test(normalizedBusinessName)) {
    errors.push("Special characters are not allowed in Business / Firm Name.");
  }

  if (!normalizedOwnerName) errors.push("Your Name is required.");
  if (!normalizedMobile) {
    errors.push("WhatsApp Number is required.");
  } else if (!/^\d{10,15}$/.test(normalizedMobile)) {
    errors.push("Enter a valid WhatsApp Number with 10 to 15 digits.");
  }

  if (!normalizedEmail) {
    errors.push("Email Address is required.");
  } else if (!registrationEmailPattern.test(normalizedEmail)) {
    errors.push("Enter a valid Email Address.");
  }

  if (!payload.password) {
    errors.push("Create Password is required.");
  } else if (String(payload.password).length < 8) {
    errors.push("Password must be at least 8 characters long.");
  }

  if (!normalizedCountry) errors.push("Country is required.");
  if (!normalizedState) errors.push("State is required.");
  if (!normalizedDistrict) errors.push("District is required.");
  if (!normalizedCity) errors.push("City is required.");
  if (!normalizedPinCode) {
    errors.push("PIN Code is required.");
  } else if (!registrationPinPattern.test(normalizedPinCode)) {
    errors.push("Invalid PIN Code.");
  }

  if (!normalizedAddress) {
    errors.push("Full Address is required.");
  } else if (normalizedAddress.length < 20) {
    errors.push("Full Address must be at least 20 characters.");
  }

  if (normalizedGstNumber && !registrationGstPattern.test(normalizedGstNumber)) {
    errors.push("Invalid GST / Tax Number format.");
  }

  if (!normalizedServices.length) {
    errors.push("Please select at least one Interested Service.");
  } else if (normalizedServices.some((service) => !allowedRegistrationServices.has(service))) {
    errors.push("One or more Interested Services are invalid.");
  }

  if (!payload.termsAccepted) {
    errors.push("Please accept the Terms & Conditions.");
  }

  if (!String(payload.captchaChallengeId || "").trim()) {
    errors.push("CAPTCHA verification is required.");
  }

  return {
    errors,
    normalized: {
      businessName: normalizedBusinessName,
      ownerName: normalizedOwnerName,
      mobile: normalizedMobile,
      email: normalizedEmail,
      password: String(payload.password || ""),
      referenceCode: normalizedReferenceCode,
      country: normalizedCountry,
      state: normalizedState,
      district: normalizedDistrict,
      city: normalizedCity,
      pinCode: normalizedPinCode,
      gstNumber: normalizedGstNumber,
      address: normalizedAddress,
      services: normalizedServices,
      termsAccepted: Boolean(payload.termsAccepted),
      captcha: String(payload.captcha || "").trim(),
      captchaChallengeId: String(payload.captchaChallengeId || "").trim()
    }
  };
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeBoolean(value) {
  if (typeof value === "boolean") return value;
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  if (["true", "1", "yes", "enabled"].includes(normalized)) return true;
  if (["false", "0", "no", "disabled"].includes(normalized)) return false;
  return null;
}

function normalizeStatusInput(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "active") return STATUS_ACTIVE;
  if (["deactive", "inactive", "disabled", "suspended"].includes(normalized)) return STATUS_SUSPENDED;
  return "";
}

function normalizeOrderStatusInput(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return "";
  if (["pending", "pending-review"].includes(normalized)) return ORDER_STATUS_PENDING;
  if (["printing", "processing", "in-printing"].includes(normalized)) return ORDER_STATUS_PRINTING;
  if (["packaging", "packing"].includes(normalized)) return ORDER_STATUS_PACKAGING;
  if (["dispatch", "dispatched", "shipment"].includes(normalized)) return ORDER_STATUS_DISPATCHED;
  if (["cancelled", "canceled"].includes(normalized)) return ORDER_STATUS_CANCELLED;
  if (["improper", "issue", "flagged"].includes(normalized)) return ORDER_STATUS_IMPROPER;
  if (["completed", "complete", "delivered"].includes(normalized)) return ORDER_STATUS_COMPLETED;
  if (["rejected", "reject"].includes(normalized)) return ORDER_STATUS_REJECTED;
  return "";
}

function normalizeOrderStatusLabel(statusValue) {
  switch (statusValue) {
    case ORDER_STATUS_PRINTING:
      return "Printing";
    case ORDER_STATUS_PACKAGING:
      return "Packaging";
    case ORDER_STATUS_DISPATCHED:
      return "Dispatched";
    case ORDER_STATUS_CANCELLED:
      return "Cancelled";
    case ORDER_STATUS_IMPROPER:
      return "Improper";
    case ORDER_STATUS_COMPLETED:
      return "Completed";
    case ORDER_STATUS_REJECTED:
      return "Rejected";
    default:
      return "Pending";
  }
}

function normalizeDesignSubmissionSourceInput(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return "";
  if (["online", "online-upload", "upload", "uploaded"].includes(normalized)) return DESIGN_SOURCE_ONLINE;
  if (["email", "mail", "send design via email"].includes(normalized)) return DESIGN_SOURCE_EMAIL;
  return "";
}

function normalizeDesignSubmissionSourceLabel(value) {
  return value === DESIGN_SOURCE_EMAIL ? "Send Design via Email" : "Online Upload";
}

function extractOrderOverviewValue(overview, label) {
  const match = String(overview || "").match(new RegExp(`${label}:\\s*([^|]+)`));
  return match ? String(match[1]).trim() : "";
}

function buildRecentOrderSummary(order) {
  const overview = order.orderDetailsOverview || "";
  const bagName = order.bagName || extractOrderOverviewValue(overview, "Product") || order.orderName || "Order";
  const quantity = order.quantity || extractOrderOverviewValue(overview, "Quantity") || "--";
  const bagSize = order.bagSize || extractOrderOverviewValue(overview, "Size") || "--";
  const bagColor = order.bagColor || extractOrderOverviewValue(overview, "Bag Color") || "--";
  return `${bagName}, Qty ${quantity}, Size ${bagSize}, Color ${bagColor}`;
}

function getRecentOrderFileType(order) {
  if (order.designSubmissionSource === DESIGN_SOURCE_EMAIL) {
    return "email";
  }

  const explicitType = String(order.designFileType || "").trim();
  if (explicitType) {
    return explicitType.toUpperCase();
  }

  const fileName = String(order.designFileName || "").trim();
  const extension = fileName.includes(".") ? fileName.split(".").pop() : "";
  return extension ? extension.toUpperCase() : "FILE";
}

function normalizeDeliveryStatusLabel(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return "Dispatch Pending";
  if (normalized === "dispatched") return "Dispatched";
  if (["in-transit", "transit"].includes(normalized)) return "In Transit";
  if (normalized === "out-for-delivery") return "Out For Delivery";
  if (normalized === "delivered") return "Delivered";
  if (["returned", "rto"].includes(normalized)) return "Returned";
  return normalized
    .split("-")
    .map((part) => (part ? part.charAt(0).toUpperCase() + part.slice(1) : ""))
    .join(" ");
}

function getOrderSortQuery(sortBy = "orderedAt", sortOrder = "desc") {
  const direction = String(sortOrder).toLowerCase() === "asc" ? 1 : -1;

  switch (sortBy) {
    case "orderNumber":
      return { orderNumber: direction, createdAt: -1 };
    case "orderName":
      return { orderName: direction, createdAt: -1 };
    case "orderDateTime":
    case "orderedAt":
      return { orderedAt: direction, createdAt: -1 };
    case "orderDetailsOverview":
      return { orderDetailsOverview: direction, createdAt: -1 };
    case "customerName":
      return { customerName: direction, createdAt: -1 };
    case "mobileNumber":
    case "customerMobile":
      return { customerMobile: direction, createdAt: -1 };
    case "status":
      return { status: direction, orderedAt: -1 };
    case "designFileSource":
      return { designSubmissionSource: direction, orderedAt: -1 };
    default:
      return { orderedAt: direction, createdAt: -1 };
  }
}

async function getNextOrderNumber() {
  const latestOrder = await Order.findOne({ orderNumber: { $ne: null } })
    .sort({ orderNumber: -1 })
    .select("orderNumber");

  if (!latestOrder?.orderNumber) {
    return 101;
  }

  return latestOrder.orderNumber + 5;
}

async function ensureOrderNumbersForExistingOrders() {
  const missingOrders = await Order.find({
    $or: [{ orderNumber: null }, { orderNumber: { $exists: false } }]
  })
    .sort({ createdAt: 1, _id: 1 })
    .select("_id");

  if (missingOrders.length === 0) {
    return;
  }

  let nextOrderNumber = await getNextOrderNumber();
  const operations = missingOrders.map((order) => {
    const operation = {
      updateOne: {
        filter: { _id: order._id },
        update: { $set: { orderNumber: nextOrderNumber } }
      }
    };
    nextOrderNumber += 5;
    return operation;
  });

  if (operations.length > 0) {
    await Order.bulkWrite(operations);
  }
}

function normalizeDesignFileUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/uploads/")) return raw;
  if (raw.startsWith("uploads/")) return `/${raw}`;
  return "";
}

function isValidDesignFileUrl(value) {
  const normalized = normalizeDesignFileUrl(value);
  return Boolean(normalized);
}

function validateCsvOrderRecord(record, rowNumber, allowedStatuses = []) {
  const errors = [];

  if (!record.orderName) errors.push(`Row ${rowNumber}: Order Name is required.`);
  if (!record.customerName) errors.push(`Row ${rowNumber}: Customer Name is required.`);
  if (!record.customerMobile) errors.push(`Row ${rowNumber}: Mobile Number is required.`);
  if (!record.orderDetailsOverview) errors.push(`Row ${rowNumber}: Order Details Overview is required.`);
  if (
    record.customerMobile &&
    !/^\d{10,15}$/.test(String(record.customerMobile).replace(/\D/g, ""))
  ) {
    errors.push(`Row ${rowNumber}: Mobile Number must contain 10 to 15 digits.`);
  }
  const normalizedStatus = normalizeOrderStatusInput(record.status);
  if (record.status && !normalizedStatus) {
    errors.push(
      `Row ${rowNumber}: Status must be Pending, Printing, Packaging, Dispatched, Cancelled, Improper, Completed, or Rejected.`
    );
  }
  if (normalizedStatus && allowedStatuses.length > 0 && !allowedStatuses.includes(normalizedStatus)) {
    errors.push(`Row ${rowNumber}: Status must be ${allowedStatuses.map(normalizeOrderStatusLabel).join(" or ")}.`);
  }
  if (record.designSubmissionSource && !normalizeDesignSubmissionSourceInput(record.designSubmissionSource)) {
    errors.push(`Row ${rowNumber}: Design Submission Source must be Online Upload or Email.`);
  }
  if (record.orderDateTime && Number.isNaN(new Date(record.orderDateTime).getTime())) {
    errors.push(`Row ${rowNumber}: Order Date & Time is invalid.`);
  }
  if (record.dispatchDateTime && Number.isNaN(new Date(record.dispatchDateTime).getTime())) {
    errors.push(`Row ${rowNumber}: Dispatch Date & Time is invalid.`);
  }
  if (record.deliveryDateTime && Number.isNaN(new Date(record.deliveryDateTime).getTime())) {
    errors.push(`Row ${rowNumber}: Delivery Date & Time is invalid.`);
  }

  return errors;
}

async function buildOrderRow(orderDocument) {
  const order = orderDocument.toObject ? orderDocument.toObject() : orderDocument;
  const placedByUser = order.placedByUserId && typeof order.placedByUserId === "object" ? order.placedByUserId : null;
  const assignedAssociate =
    order.assignedAssociateMemberId && typeof order.assignedAssociateMemberId === "object" ? order.assignedAssociateMemberId : null;
  const assignedAdmin = order.assignedAdminId && typeof order.assignedAdminId === "object" ? order.assignedAdminId : null;

  return {
    id: order._id.toString(),
    orderNumber: order.orderNumber,
    orderName: order.orderName || order.referenceNo || `Order ${order.orderNumber || ""}`.trim(),
    orderDateTime: order.orderedAt || order.createdAt,
    orderDetailsOverview: order.orderDetailsOverview || order.referenceNo || "No overview available",
    currentStatus: normalizeOrderStatusLabel(order.status),
    currentStatusValue: order.status,
    designFileSource: normalizeDesignSubmissionSourceLabel(order.designSubmissionSource),
    designFileSourceValue: order.designSubmissionSource,
    hasFileAttachment:
      order.designSubmissionSource === DESIGN_SOURCE_ONLINE && Boolean(order.designFileName) && isValidDesignFileUrl(order.designFileUrl),
    hasEmailDesign: order.designSubmissionSource === DESIGN_SOURCE_EMAIL,
    designFileName: order.designFileName || "",
    designFileType: order.designFileType || "",
    designFileUrl: normalizeDesignFileUrl(order.designFileUrl),
    courierName: order.courierName || "",
    courierTrackingNumber: order.courierTrackingNumber || "",
    courierTrackingUrl: order.courierTrackingUrl || "",
    dispatchDateTime: order.dispatchDateTime || null,
    dispatchNotes: order.dispatchNotes || "",
    deliveryStatus: normalizeDeliveryStatusLabel(order.deliveryStatus),
    deliveryStatusValue: order.deliveryStatus || "",
    deliveryDateTime: order.deliveryDateTime || null,
    customerName: order.customerName || placedByUser?.ownerName || "",
    mobileNumber: order.customerMobile || placedByUser?.mobile || "",
    bagName: order.bagName || "",
    printSide: order.printSide || "",
    quantity: Number(order.quantity || 0),
    bagSize: order.bagSize || "",
    bagColor: order.bagColor || "",
    textColorType: order.textColorType || "",
    textColors: Array.isArray(order.textColors) ? order.textColors : [],
    printingPress: order.printingPress || "",
    privacy: order.privacy || "",
    deliveryOption: order.deliveryOption || "",
    fileOption: order.fileOption || "",
    sellingPrice: Number(order.sellingPrice || 0),
    remark: order.remark || "",
    pressline: order.pressline || "",
    placedByUser: placedByUser
      ? {
          id: placedByUser._id?.toString?.() || "",
          name: placedByUser.ownerName || "",
          mobileNumber: placedByUser.mobile || "",
          businessName: placedByUser.businessName || "",
          email: placedByUser.email || "",
          country: placedByUser.country || "",
          state: placedByUser.state || "",
          district: placedByUser.district || "",
          city: placedByUser.city || "",
          pinCode: placedByUser.pinCode || "",
          address: placedByUser.address || "",
          fullAddress: buildAddress(placedByUser),
          gstNumber: placedByUser.gstNumber || "",
          referenceCode: placedByUser.referenceCode || "",
          registrationDate: placedByUser.createdAt || null,
          role: placedByUser.role || "",
          associateMemberId: getAssociateMemberDisplayId(placedByUser.mobile || ""),
          status: normalizeAdminStatusLabel(placedByUser.status)
        }
      : null,
    assignedAssociateMember: assignedAssociate
      ? {
          id: assignedAssociate._id?.toString?.() || "",
          name: assignedAssociate.ownerName || "",
          mobileNumber: assignedAssociate.mobile || "",
          businessName: assignedAssociate.businessName || "",
          email: assignedAssociate.email || "",
          status: normalizeAdminStatusLabel(assignedAssociate.status)
        }
      : null,
    assignedAdmin: assignedAdmin
      ? {
          id: assignedAdmin._id?.toString?.() || "",
          name: assignedAdmin.ownerName || "",
          mobileNumber: assignedAdmin.mobile || "",
          businessName: assignedAdmin.businessName || "",
          email: assignedAdmin.email || "",
          status: normalizeAdminStatusLabel(assignedAdmin.status)
        }
      : null,
    referenceNo: order.referenceNo || "",
    basePayableAmount: Number(order.basePayableAmount || 0),
    pdfDiscountAmount: Number(order.pdfDiscountAmount || 0),
    walletDebitAmount: Number(order.walletDebitAmount || 0),
    isUrgent: Boolean(order.isUrgent),
    createdAt: order.createdAt,
    updatedAt: order.updatedAt
  };
}

async function buildOrderDetails(orderDocument) {
  const populatedOrder = await Order.findById(orderDocument._id)
    .populate("placedByUserId", "ownerName mobile email businessName country state district city pinCode address gstNumber referenceCode createdAt status role")
    .populate("assignedAssociateMemberId", "ownerName mobile email businessName status")
    .populate("assignedAdminId", "ownerName mobile email businessName status")
    .select(
      "orderNumber orderName customerName customerMobile orderDetailsOverview bagName printSide quantity bagSize bagColor textColorType textColors printingPress privacy deliveryOption fileOption sellingPrice remark pressline orderedAt status designSubmissionSource designFileName designFileType designFileUrl courierName courierTrackingNumber courierTrackingUrl dispatchDateTime dispatchNotes deliveryStatus deliveryDateTime isUrgent referenceNo basePayableAmount pdfDiscountAmount walletDebitAmount placedByUserId assignedAssociateMemberId assignedAdminId statusHistory dispatchHistory createdAt updatedAt"
    );

  const baseRow = await buildOrderRow(populatedOrder);

  return {
    ...baseRow,
    statusHistory: (populatedOrder.statusHistory || [])
      .map((entry, index) => ({
        id: `${populatedOrder._id.toString()}-${index}-${entry.changedAt?.toString?.() || "history"}`,
        status: normalizeOrderStatusLabel(entry.status),
        statusValue: entry.status,
        note: entry.note || "",
        changedAt: entry.changedAt,
        changedByRole: entry.changedByRole || "",
        changedById: entry.changedById || "",
        changedByName: entry.changedByName || "System"
      }))
      .sort((left, right) => new Date(right.changedAt).getTime() - new Date(left.changedAt).getTime()),
    dispatchHistory: (populatedOrder.dispatchHistory || [])
      .map((entry, index) => ({
        id: `${populatedOrder._id.toString()}-dispatch-${index}-${entry.eventAt?.toString?.() || "history"}`,
        status: normalizeOrderStatusLabel(entry.status),
        statusValue: entry.status,
        note: entry.note || "",
        courierName: entry.courierName || "",
        courierTrackingNumber: entry.courierTrackingNumber || "",
        courierTrackingUrl: entry.courierTrackingUrl || "",
        deliveryStatus: normalizeDeliveryStatusLabel(entry.deliveryStatus),
        deliveryStatusValue: entry.deliveryStatus || "",
        eventAt: entry.eventAt,
        updatedByRole: entry.updatedByRole || "",
        updatedById: entry.updatedById || "",
        updatedByName: entry.updatedByName || "System"
      }))
      .sort((left, right) => new Date(right.eventAt).getTime() - new Date(left.eventAt).getTime())
  };
}

async function fetchOrderList(queryParams = {}) {
  await ensureOrderNumbersForExistingOrders();

  const {
    placedByUserId = "",
    search = "",
    orderNumber = "",
    orderName = "",
    customerName = "",
    mobileNumber = "",
    status = "",
    designFileSource = "",
    fromDate = "",
    toDate = "",
    allowedStatuses = [],
    page = 1,
    limit = 10,
    sortBy = "orderedAt",
    sortOrder = "desc"
  } = queryParams;

  const query = {};
  const andFilters = [];
  if (placedByUserId) {
    query.placedByUserId = placedByUserId;
  }
  const normalizedAllowedStatuses = (Array.isArray(allowedStatuses) ? allowedStatuses : [allowedStatuses])
    .map((statusValue) => normalizeOrderStatusInput(statusValue))
    .filter(Boolean);

  const trimmedSearch = String(search).trim();
  if (trimmedSearch) {
    const searchRegex = new RegExp(escapeRegex(trimmedSearch), "i");
    const digits = trimmedSearch.replace(/\D/g, "");
    const searchConditions = [
      { orderName: searchRegex },
      { customerName: searchRegex },
      { orderDetailsOverview: searchRegex },
      { referenceNo: searchRegex }
    ];

    if (digits) {
      searchConditions.push({ customerMobile: new RegExp(escapeRegex(digits)) });
      if (!Number.isNaN(Number(digits))) {
        searchConditions.push({ orderNumber: Number(digits) });
      }
    }

    const normalizedSearchStatus = normalizeOrderStatusInput(trimmedSearch);
    if (normalizedSearchStatus) {
      searchConditions.push({ status: normalizedSearchStatus });
    }

    andFilters.push({ $or: searchConditions });
  }

  if (orderNumber) {
    const numericOrderNumber = Number(String(orderNumber).replace(/\D/g, ""));
    if (!Number.isNaN(numericOrderNumber) && numericOrderNumber > 0) {
      andFilters.push({ orderNumber: numericOrderNumber });
    }
  }

  if (orderName) {
    andFilters.push({ orderName: new RegExp(escapeRegex(String(orderName).trim()), "i") });
  }

  if (customerName) {
    andFilters.push({ customerName: new RegExp(escapeRegex(String(customerName).trim()), "i") });
  }

  if (mobileNumber) {
    andFilters.push({ customerMobile: new RegExp(escapeRegex(String(mobileNumber).replace(/\D/g, ""))) });
  }

  const normalizedOrderStatus = normalizeOrderStatusInput(status);
  if (normalizedOrderStatus) {
    query.status =
      normalizedAllowedStatuses.length > 0 && !normalizedAllowedStatuses.includes(normalizedOrderStatus)
        ? { $in: [] }
        : normalizedOrderStatus;
  } else if (normalizedAllowedStatuses.length > 0) {
    query.status = { $in: normalizedAllowedStatuses };
  }

  const normalizedDesignSource = normalizeDesignSubmissionSourceInput(designFileSource);
  if (normalizedDesignSource) {
    query.designSubmissionSource = normalizedDesignSource;
  }

  if (fromDate || toDate) {
    query.orderedAt = {};
    if (fromDate) {
      query.orderedAt.$gte = new Date(`${fromDate}T00:00:00.000Z`);
    }
    if (toDate) {
      query.orderedAt.$lte = new Date(`${toDate}T23:59:59.999Z`);
    }
  }

  if (andFilters.length > 0) {
    query.$and = andFilters;
  }

  const safePage = Math.max(Number(page) || 1, 1);
  const safeLimit = Math.min(Math.max(Number(limit) || 10, 1), 100);
  const sortQuery = getOrderSortQuery(sortBy, sortOrder);

  const [orders, totalRecords] = await Promise.all([
    Order.find(query)
      .populate("placedByUserId", "ownerName mobile email businessName")
      .populate("assignedAssociateMemberId", "ownerName mobile email businessName status")
      .populate("assignedAdminId", "ownerName mobile email businessName status")
      .sort(sortQuery)
      .skip((safePage - 1) * safeLimit)
      .limit(safeLimit),
    Order.countDocuments(query)
  ]);

  const items = await Promise.all(orders.map((order) => buildOrderRow(order)));

  return {
    items,
    pagination: {
      page: safePage,
      limit: safeLimit,
      totalRecords,
      totalPages: Math.max(Math.ceil(totalRecords / safeLimit), 1)
    }
  };
}

function mapAssociateSearchOrderRow(order) {
  return {
    id: order.id,
    orderNumber: order.orderNumber || "--",
    dateTime: formatDateTime(order.orderDateTime || order.createdAt),
    orderName: order.orderName || "--",
    orderDetail: buildRecentOrderSummary(order),
    status: order.currentStatus || "--",
    currentStatusValue: order.currentStatusValue || "",
    fileType: getRecentOrderFileType(order),
    bagName: order.bagName || "--",
    printSide: order.printSide || "--",
    quantity: Number(order.quantity || 0),
    bagSize: order.bagSize || "--",
    bagColor: order.bagColor || "--",
    textColorType: order.textColorType || "--",
    textColors: Array.isArray(order.textColors) ? order.textColors : [],
    customerName: order.customerName || "--",
    customerMobile: order.mobileNumber || "--",
    orderDetailsOverview: order.orderDetailsOverview || "--",
    printingPress: order.printingPress || "--",
    privacy: order.privacy || "--",
    deliveryOption: order.deliveryOption || "--",
    fileOption: order.fileOption || "--",
    sellingPrice: Number(order.sellingPrice || 0),
    remark: order.remark || "",
    pressline: order.pressline || "--",
    referenceNo: order.referenceNo || "--",
    designSubmissionSource: order.designSubmissionSourceValue || "",
    designFileName: order.designFileName || "",
    designFileType: order.designFileType || "",
    designFileUrl: order.designFileUrl || "",
    basePayableAmount: Number(order.basePayableAmount || 0),
    pdfDiscountAmount: Number(order.pdfDiscountAmount || 0),
    walletDebitAmount: Number(order.walletDebitAmount || 0)
  };
}

function mapRecentOrderRow(order) {
  return {
    id: order._id.toString(),
    orderNumber: order.orderNumber || "--",
    dateTime: formatDateTime(order.orderedAt || order.createdAt),
    orderName: order.orderName || "--",
    orderDetail: buildRecentOrderSummary(order),
    status: normalizeOrderStatusLabel(order.status),
    currentStatusValue: order.status || "",
    fileType: getRecentOrderFileType(order),
    bagName: order.bagName || "--",
    printSide: order.printSide || "--",
    quantity: Number(order.quantity || 0),
    bagSize: order.bagSize || "--",
    bagColor: order.bagColor || "--",
    textColorType: order.textColorType || "--",
    textColors: Array.isArray(order.textColors) ? order.textColors : [],
    customerName: order.customerName || "--",
    customerMobile: order.customerMobile || "--",
    orderDetailsOverview: order.orderDetailsOverview || "--",
    printingPress: order.printingPress || "--",
    privacy: order.privacy || "--",
    deliveryOption: order.deliveryOption || "--",
    fileOption: order.fileOption || "--",
    sellingPrice: Number(order.sellingPrice || 0),
    remark: order.remark || "",
    pressline: order.pressline || "--",
    referenceNo: order.referenceNo || "--",
    designSubmissionSource: order.designSubmissionSource || "",
    designFileName: order.designFileName || "",
    designFileType: order.designFileType || "",
    designFileUrl: order.designFileUrl || "",
    basePayableAmount: Number(order.basePayableAmount || 0),
    pdfDiscountAmount: Number(order.pdfDiscountAmount || 0),
    walletDebitAmount: Number(order.walletDebitAmount || 0)
  };
}

function getAdminSortQuery(sortBy = "createdAt", sortOrder = "desc") {
  const direction = String(sortOrder).toLowerCase() === "asc" ? 1 : -1;

  switch (sortBy) {
    case "adminId":
      return { mobile: direction, createdAt: -1 };
    case "name":
      return { ownerName: direction, createdAt: -1 };
    case "mobile":
      return { mobile: direction, createdAt: -1 };
    case "address":
      return { city: direction, state: direction, createdAt: -1 };
    case "businessName":
      return { businessName: direction, createdAt: -1 };
    case "status":
      return { status: direction, createdAt: -1 };
    case "associateAccess":
      return { associateMemberAccessEnabled: direction, createdAt: -1 };
    default:
      return { createdAt: direction };
  }
}

function parseCsvLine(line) {
  const values = [];
  let current = "";
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"' && insideQuotes && nextChar === '"') {
      current += '"';
      index += 1;
      continue;
    }

    if (char === '"') {
      insideQuotes = !insideQuotes;
      continue;
    }

    if (char === "," && !insideQuotes) {
      values.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  values.push(current.trim());
  return values;
}

function parseCsvText(csvText) {
  return String(csvText)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => parseCsvLine(line));
}

function validateCsvAdminRecord(record, rowNumber) {
  const errors = [];

  if (!record.ownerName) errors.push(`Row ${rowNumber}: Admin Name is required.`);
  if (!record.mobile) errors.push(`Row ${rowNumber}: Mobile Number is required.`);
  if (!record.email) errors.push(`Row ${rowNumber}: Email is required.`);
  if (!record.businessName) errors.push(`Row ${rowNumber}: Business/Firm Name is required.`);
  if (!record.address) errors.push(`Row ${rowNumber}: Admin Address is required.`);
  if (!record.country) errors.push(`Row ${rowNumber}: Country is required.`);
  if (!record.state) errors.push(`Row ${rowNumber}: State is required.`);
  if (!record.district) errors.push(`Row ${rowNumber}: District is required.`);
  if (!record.city) errors.push(`Row ${rowNumber}: City is required.`);
  if (!record.pinCode) errors.push(`Row ${rowNumber}: Pin Code is required.`);
  if (!record.password) errors.push(`Row ${rowNumber}: Password is required.`);
  if (record.email && !String(record.email).includes("@")) errors.push(`Row ${rowNumber}: Email format is invalid.`);
  if (record.mobile && !/^\d{10,15}$/.test(String(record.mobile).replace(/\D/g, ""))) {
    errors.push(`Row ${rowNumber}: Mobile Number must contain 10 to 15 digits.`);
  }
  if (record.status && !["active", "deactive", "inactive", "disabled", "suspended"].includes(String(record.status).trim().toLowerCase())) {
    errors.push(`Row ${rowNumber}: Status must be Active or Deactive.`);
  }
  if (record.associateMemberAccessStatus) {
    const parsed = normalizeBoolean(record.associateMemberAccessStatus);
    if (parsed === null) {
      errors.push(`Row ${rowNumber}: Associate Member Access Status must be Enabled or Disabled.`);
    }
  }

  return errors;
}

async function createAuditLog({
  actorRole,
  actorId,
  actorName,
  module,
  action,
  targetType,
  targetId,
  targetName,
  details
}) {
  if (mongoose.connection.readyState !== 1) return;

  await AuditLog.create({
    actorRole,
    actorId,
    actorName,
    module,
    action,
    targetType,
    targetId,
    targetName,
    details
  });
}

async function getActorName(auth) {
  if (!auth) return "System";
  if (auth.role === ROLE_SUPER_ADMIN) return "Super Admin";
  if (!mongoose.Types.ObjectId.isValid(auth.sub)) return "User";
  const actor = await User.findById(auth.sub).select("ownerName");
  return actor?.ownerName || "User";
}

async function buildAdminRow(adminDocument, associateMemberCount) {
  return {
    id: adminDocument._id.toString(),
    adminId: getAdminDisplayId(adminDocument.mobile),
    adminName: adminDocument.ownerName,
    mobileNumber: adminDocument.mobile,
    adminAddress: buildAddress(adminDocument),
    businessName: adminDocument.businessName,
    status: normalizeAdminStatusLabel(adminDocument.status),
    statusValue: adminDocument.status,
    associateMemberAccessStatus: adminDocument.associateMemberAccessEnabled ? "Enabled" : "Disabled",
    associateMemberAccessEnabled: Boolean(adminDocument.associateMemberAccessEnabled),
    associateMemberCount,
    email: adminDocument.email,
    country: adminDocument.country,
    state: adminDocument.state,
    district: adminDocument.district,
    city: adminDocument.city,
    pinCode: adminDocument.pinCode,
    gstNumber: adminDocument.gstNumber || "",
    referenceCode: adminDocument.referenceCode || "",
    createdAt: adminDocument.createdAt,
    updatedAt: adminDocument.updatedAt
  };
}

async function buildAdminDetails(adminDocument) {
  const associateMembers = await User.find({
    role: ROLE_ASSOCIATE_MEMBER,
    assignedAdminId: adminDocument._id
  })
    .sort({ ownerName: 1 })
    .select("ownerName businessName mobile city state status");

  const baseRow = await buildAdminRow(adminDocument, associateMembers.length);

  return {
    ...baseRow,
    associateMembers: associateMembers.map((member) => ({
      id: member._id.toString(),
      ownerName: member.ownerName,
      businessName: member.businessName,
      mobileNumber: member.mobile,
      location: [member.city, member.state].filter(Boolean).join(", "),
      status: normalizeAdminStatusLabel(member.status)
    }))
  };
}

async function syncAssociateMemberOrdersAdminAssignment(associateMemberId, assignedAdminId) {
  if (!associateMemberId) return;

  await Order.updateMany(
    {
      $or: [{ placedByUserId: associateMemberId }, { assignedAssociateMemberId: associateMemberId }]
    },
    {
      $set: {
        assignedAdminId: assignedAdminId || null
      }
    }
  );
}

async function fetchAdminList(queryParams = {}) {
  const {
    search = "",
    name = "",
    mobile = "",
    adminId = "",
    businessName = "",
    status = "",
    associateAccess = "",
    page = 1,
    limit = 10,
    sortBy = "createdAt",
    sortOrder = "desc"
  } = queryParams;

  const query = { role: ROLE_ADMIN };
  const andFilters = [];

  const trimmedSearch = String(search).trim();
  if (trimmedSearch) {
    const searchRegex = new RegExp(escapeRegex(trimmedSearch), "i");
    const digits = trimmedSearch.replace(/\D/g, "");
    const searchConditions = [
      { ownerName: searchRegex },
      { businessName: searchRegex },
      { email: searchRegex }
    ];

    if (digits) {
      searchConditions.push({ mobile: new RegExp(escapeRegex(digits)) });
    }

    const adminIdDigits = trimmedSearch.toUpperCase().startsWith("ADM-") ? trimmedSearch.slice(4).replace(/\D/g, "") : digits;
    if (adminIdDigits) {
      searchConditions.push({ mobile: new RegExp(`${escapeRegex(adminIdDigits)}$`) });
    }

    andFilters.push({ $or: searchConditions });
  }

  if (name) {
    andFilters.push({ ownerName: new RegExp(escapeRegex(String(name).trim()), "i") });
  }

  if (mobile) {
    andFilters.push({ mobile: new RegExp(escapeRegex(String(mobile).replace(/\D/g, ""))) });
  }

  if (adminId) {
    const adminDigits = String(adminId).replace(/\D/g, "");
    if (adminDigits) {
      andFilters.push({ mobile: new RegExp(`${escapeRegex(adminDigits)}$`) });
    }
  }

  if (businessName) {
    andFilters.push({ businessName: new RegExp(escapeRegex(String(businessName).trim()), "i") });
  }

  const normalizedStatus = normalizeStatusInput(status);
  if (normalizedStatus) {
    query.status = normalizedStatus;
  }

  const normalizedAssociateAccess = normalizeBoolean(associateAccess);
  if (normalizedAssociateAccess !== null) {
    query.associateMemberAccessEnabled = normalizedAssociateAccess;
  }

  if (andFilters.length > 0) {
    query.$and = andFilters;
  }

  const safePage = Math.max(Number(page) || 1, 1);
  const safeLimit = Math.min(Math.max(Number(limit) || 10, 1), 100);
  const sortQuery = getAdminSortQuery(sortBy, sortOrder);

  const [admins, totalRecords] = await Promise.all([
    User.find(query)
      .sort(sortQuery)
      .skip((safePage - 1) * safeLimit)
      .limit(safeLimit),
    User.countDocuments(query)
  ]);

  const associateCounts = await User.aggregate([
    {
      $match: {
        role: ROLE_ASSOCIATE_MEMBER,
        assignedAdminId: { $in: admins.map((admin) => admin._id) }
      }
    },
    {
      $group: {
        _id: "$assignedAdminId",
        count: { $sum: 1 }
      }
    }
  ]);

  const countMap = new Map(associateCounts.map((entry) => [String(entry._id), entry.count]));
  const items = await Promise.all(admins.map((admin) => buildAdminRow(admin, countMap.get(String(admin._id)) || 0)));

  return {
    items,
    pagination: {
      page: safePage,
      limit: safeLimit,
      totalRecords,
      totalPages: Math.max(Math.ceil(totalRecords / safeLimit), 1)
    }
  };
}

function getAssociateMemberDisplayId(mobileNumber) {
  const mobile = String(mobileNumber || "").trim();
  const digits = mobile.replace(/\D/g, "");
  return `ASM-${digits.slice(-4).padStart(4, "0")}`;
}

function getAssociateMemberSortQuery(sortBy = "createdAt", sortOrder = "desc") {
  const direction = String(sortOrder).toLowerCase() === "asc" ? 1 : -1;

  switch (sortBy) {
    case "associateMemberId":
      return { mobile: direction, createdAt: -1 };
    case "businessName":
      return { businessName: direction, createdAt: -1 };
    case "name":
      return { ownerName: direction, createdAt: -1 };
    case "mobile":
      return { mobile: direction, createdAt: -1 };
    case "district":
      return { district: direction, createdAt: -1 };
    case "assignedAdminName":
      return { assignedAdminName: direction, ownerName: 1 };
    case "status":
      return { status: direction, createdAt: -1 };
    default:
      return { createdAt: direction };
  }
}

function validateCsvAssociateMemberRecord(record, rowNumber) {
  const errors = [];

  if (!record.ownerName) errors.push(`Row ${rowNumber}: Associate Member Name is required.`);
  if (!record.mobile) errors.push(`Row ${rowNumber}: Mobile Number is required.`);
  if (!record.email) errors.push(`Row ${rowNumber}: Email is required.`);
  if (!record.businessName) errors.push(`Row ${rowNumber}: Business/Firm Name is required.`);
  if (!record.address) errors.push(`Row ${rowNumber}: Address is required.`);
  if (!record.country) errors.push(`Row ${rowNumber}: Country is required.`);
  if (!record.state) errors.push(`Row ${rowNumber}: State is required.`);
  if (!record.district) errors.push(`Row ${rowNumber}: District is required.`);
  if (!record.city) errors.push(`Row ${rowNumber}: City is required.`);
  if (!record.pinCode) errors.push(`Row ${rowNumber}: Pin Code is required.`);
  if (!record.password) errors.push(`Row ${rowNumber}: Password is required.`);
  if (record.email && !String(record.email).includes("@")) errors.push(`Row ${rowNumber}: Email format is invalid.`);
  if (record.mobile && !/^\d{10,15}$/.test(String(record.mobile).replace(/\D/g, ""))) {
    errors.push(`Row ${rowNumber}: Mobile Number must contain 10 to 15 digits.`);
  }
  if (record.status && !["active", "deactive", "inactive", "disabled", "suspended", "pending-review", "pending"].includes(String(record.status).trim().toLowerCase())) {
    errors.push(`Row ${rowNumber}: Status must be Active or Deactive.`);
  }

  return errors;
}

async function buildAssociateMemberRow(memberDocument, assignedAdminDocument = null) {
  const assignedAdmin = assignedAdminDocument || (memberDocument.assignedAdminId && typeof memberDocument.assignedAdminId === "object"
    ? memberDocument.assignedAdminId
    : null);

  return {
    id: memberDocument._id.toString(),
    associateMemberId: getAssociateMemberDisplayId(memberDocument.mobile),
    associateMemberName: memberDocument.ownerName,
    mobileNumber: memberDocument.mobile,
    assignedAdminId: assignedAdmin?._id?.toString?.() || "",
    assignedAdminName: assignedAdmin?.ownerName || "Not Assigned",
    assignedAdminMobileNumber: assignedAdmin?.mobile || "",
    assignedAdminStatus: assignedAdmin ? normalizeAdminStatusLabel(assignedAdmin.status) : "",
    businessName: memberDocument.businessName,
    address: buildAddress(memberDocument),
    email: memberDocument.email,
    country: memberDocument.country,
    state: memberDocument.state,
    district: memberDocument.district,
    city: memberDocument.city,
    pinCode: memberDocument.pinCode,
    gstNumber: memberDocument.gstNumber || "",
    referenceCode: memberDocument.referenceCode || "",
    status: normalizeAdminStatusLabel(memberDocument.status),
    statusValue: memberDocument.status,
    role: memberDocument.role,
    createdAt: memberDocument.createdAt,
    updatedAt: memberDocument.updatedAt
  };
}

async function buildAssociateMemberDetails(memberDocument) {
  const populatedMember = await User.findById(memberDocument._id)
    .populate("assignedAdminId", "ownerName mobile email status businessName")
    .select(
      "ownerName mobile email businessName address country state district city pinCode gstNumber referenceCode status role assignedAdminId createdAt updatedAt"
    );

  const baseRow = await buildAssociateMemberRow(populatedMember, populatedMember.assignedAdminId);

  return {
    ...baseRow,
    assignedAdmin:
      populatedMember.assignedAdminId && typeof populatedMember.assignedAdminId === "object"
        ? {
            id: populatedMember.assignedAdminId._id.toString(),
            adminName: populatedMember.assignedAdminId.ownerName,
            mobileNumber: populatedMember.assignedAdminId.mobile,
            email: populatedMember.assignedAdminId.email,
            businessName: populatedMember.assignedAdminId.businessName,
            status: normalizeAdminStatusLabel(populatedMember.assignedAdminId.status)
          }
        : null
  };
}

async function fetchAssociateMemberList(queryParams = {}) {
  const {
    search = "",
    businessName = "",
    name = "",
    mobile = "",
    associateMemberId = "",
    district = "",
    assignedAdminName = "",
    status = "",
    page = 1,
    limit = 10,
    sortBy = "createdAt",
    sortOrder = "desc"
  } = queryParams;

  const query = { role: ROLE_ASSOCIATE_MEMBER };
  const andFilters = [];

  const addAssignedAdminNameFilter = async (value) => {
    const matchingAdmins = await User.find({
      role: ROLE_ADMIN,
      ownerName: new RegExp(escapeRegex(String(value).trim()), "i")
    }).select("_id");

    if (matchingAdmins.length === 0) {
      andFilters.push({ assignedAdminId: { $in: [] } });
      return;
    }

    andFilters.push({ assignedAdminId: { $in: matchingAdmins.map((admin) => admin._id) } });
  };

  const trimmedSearch = String(search).trim();
  if (trimmedSearch) {
    const searchRegex = new RegExp(escapeRegex(trimmedSearch), "i");
    const digits = trimmedSearch.replace(/\D/g, "");
    const searchConditions = [
      { ownerName: searchRegex },
      { businessName: searchRegex },
      { district: searchRegex },
      { email: searchRegex }
    ];
    const matchingAdmins = await User.find({
      role: ROLE_ADMIN,
      ownerName: new RegExp(escapeRegex(trimmedSearch), "i")
    }).select("_id");

    if (digits) {
      searchConditions.push({ mobile: new RegExp(escapeRegex(digits)) });
    }

    const associateDigits = trimmedSearch.toUpperCase().startsWith("ASM-") ? trimmedSearch.slice(4).replace(/\D/g, "") : digits;
    if (associateDigits) {
      searchConditions.push({ mobile: new RegExp(`${escapeRegex(associateDigits)}$`) });
    }

    if (matchingAdmins.length > 0) {
      searchConditions.push({ assignedAdminId: { $in: matchingAdmins.map((admin) => admin._id) } });
    }

    andFilters.push({ $or: searchConditions });
  }

  if (name) {
    andFilters.push({ ownerName: new RegExp(escapeRegex(String(name).trim()), "i") });
  }

  if (businessName) {
    andFilters.push({ businessName: new RegExp(escapeRegex(String(businessName).trim()), "i") });
  }

  if (mobile) {
    andFilters.push({ mobile: new RegExp(escapeRegex(String(mobile).replace(/\D/g, ""))) });
  }

  if (associateMemberId) {
    const associateDigits = String(associateMemberId).replace(/\D/g, "");
    if (associateDigits) {
      andFilters.push({ mobile: new RegExp(`${escapeRegex(associateDigits)}$`) });
    }
  }

  if (district) {
    andFilters.push({ district: new RegExp(escapeRegex(String(district).trim()), "i") });
  }

  if (assignedAdminName) {
    await addAssignedAdminNameFilter(assignedAdminName);
  }

  const normalizedStatus = normalizeStatusInput(status);
  if (normalizedStatus) {
    query.status = normalizedStatus;
  }

  if (andFilters.length > 0) {
    query.$and = andFilters;
  }

  const safePage = Math.max(Number(page) || 1, 1);
  const safeLimit = Math.min(Math.max(Number(limit) || 10, 1), 100);
  const sortQuery = getAssociateMemberSortQuery(sortBy, sortOrder);

  let members = [];
  let totalRecords = 0;

  if (sortBy === "assignedAdminName") {
    const allMembers = await User.find(query)
      .populate("assignedAdminId", "ownerName mobile status businessName")
      .select(
        "ownerName mobile email businessName address country state district city pinCode gstNumber referenceCode status role assignedAdminId createdAt updatedAt"
      );

    allMembers.sort((left, right) => {
      const leftValue = left.assignedAdminId?.ownerName || "";
      const rightValue = right.assignedAdminId?.ownerName || "";
      const compareResult = leftValue.localeCompare(rightValue, undefined, { sensitivity: "base" });
      return String(sortOrder).toLowerCase() === "asc" ? compareResult : compareResult * -1;
    });

    totalRecords = allMembers.length;
    members = allMembers.slice((safePage - 1) * safeLimit, safePage * safeLimit);
  } else {
    const [foundMembers, foundTotal] = await Promise.all([
      User.find(query)
        .sort(sortQuery)
        .skip((safePage - 1) * safeLimit)
        .limit(safeLimit)
        .populate("assignedAdminId", "ownerName mobile status businessName"),
      User.countDocuments(query)
    ]);

    members = foundMembers;
    totalRecords = foundTotal;
  }

  const items = await Promise.all(
    members.map((member) => buildAssociateMemberRow(member, member.assignedAdminId && typeof member.assignedAdminId === "object" ? member.assignedAdminId : null))
  );

  return {
    items,
    pagination: {
      page: safePage,
      limit: safeLimit,
      totalRecords,
      totalPages: Math.max(Math.ceil(totalRecords / safeLimit), 1)
    }
  };
}

app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", service: "auth-api" });
});

app.get("/api/auth/captcha", (_req, res) => {
  return res.json(createCaptchaChallenge());
});

app.post("/api/auth/login", async (req, res) => {
  try {
    const { country, mobileNumber, password } = req.body;

    if (!country || !mobileNumber || !password) {
      return res.status(400).json({ message: "Country, WhatsApp Number, and Password are required." });
    }

    const normalizedCountry = String(country).trim();
    const normalizedMobileNumber = normalizeMobileNumber(normalizedCountry, mobileNumber);

    if (!normalizedMobileNumber || !/^\d{10,15}$/.test(normalizedMobileNumber)) {
      return res.status(400).json({ message: "Enter a valid WhatsApp Number with 10 to 15 digits." });
    }

    const normalizedSuperAdminMobile = normalizeMobileNumber(superAdminCredentials.country, superAdminCredentials.mobileNumber);
    if (
      normalizedCountry === superAdminCredentials.country &&
      normalizedMobileNumber === normalizedSuperAdminMobile &&
      password === superAdminCredentials.password
    ) {
      const token = createToken({
        sub: "super-admin",
        country: superAdminCredentials.country,
        mobileNumber: normalizedSuperAdminMobile,
        role: superAdminCredentials.role
      });

      const walletBalance = await getWalletBalanceForActor({ role: ROLE_SUPER_ADMIN, externalKey: "super-admin" });

      return res.json({
        message: "Login successful.",
        token,
        user: buildUserResponse({
          ...superAdminCredentials,
          mobileNumber: normalizedSuperAdminMobile,
          walletBalance
        })
      });
    }

    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ message: "Database is not connected." });
    }

    const user = await User.findOne({
      country: normalizedCountry,
      mobile: normalizedMobileNumber
    });

    if (!user) {
      return res.status(401).json({ message: "Invalid Country + WhatsApp Number or Password." });
    }

    const isPasswordValid = await bcrypt.compare(String(password), user.passwordHash);
    if (!isPasswordValid) {
      return res.status(401).json({ message: "Invalid Country + WhatsApp Number or Password." });
    }

    if (user.status === STATUS_SUSPENDED) {
      return res.status(403).json({ message: "Your account has been suspended. Please contact support." });
    }

    await User.updateOne(
      { _id: user._id },
      { $set: { lastLoginAt: new Date(), lastActivityAt: new Date() } }
    );

    const token = createToken({
      sub: user._id.toString(),
      country: user.country,
      mobileNumber: user.mobile,
      role: user.role
    });

    const walletBalance = await getWalletBalanceForActor({ userId: user._id, role: user.role });

    return res.json({
      message: "Login successful.",
      token,
      user: buildUserResponse({
        ...user.toObject(),
        mobileNumber: user.mobile,
        walletBalance
      })
    });
  } catch (error) {
    return res.status(500).json({ message: error?.message || "Login failed due to a server error." });
  }
});

async function registerHandler(req, res) {
  try {
    if (mongoose.connection.readyState !== 1) {
      return res.status(503).json({ message: "Database is not connected." });
    }

    const { errors, normalized } = validateRegistrationPayload(req.body);
    if (errors.length > 0) {
      return res.status(400).json({
        message: errors[0],
        errors
      });
    }

    const captchaVerification = verifyCaptchaChallenge(normalized.captchaChallengeId, normalized.captcha);
    if (!captchaVerification.ok) {
      return res.status(400).json({ message: captchaVerification.message });
    }

    const [emailUser, mobileUser] = await Promise.all([
      User.findOne({ email: normalized.email }).select("_id"),
      User.findOne({ country: normalized.country, mobile: normalized.mobile }).select("_id")
    ]);

    if (emailUser) {
      return res.status(409).json({ message: "An account with this Email Address already exists." });
    }

    if (mobileUser) {
      return res.status(409).json({ message: "An account with this Country + WhatsApp Number already exists." });
    }

    const passwordHash = await bcrypt.hash(normalized.password, 12);

    const registration = await User.create({
      businessName: normalized.businessName,
      ownerName: normalized.ownerName,
      mobile: normalized.mobile,
      email: normalized.email,
      passwordHash,
      referenceCode: normalized.referenceCode,
      country: normalized.country,
      state: normalized.state,
      district: normalized.district,
      city: normalized.city,
      pinCode: normalized.pinCode,
      gstNumber: normalized.gstNumber,
      address: normalized.address,
      services: normalized.services,
      termsAccepted: normalized.termsAccepted,
      captchaVerified: true,
      role: ROLE_ASSOCIATE_MEMBER,
      status: STATUS_PENDING
    });

    return res.status(201).json({
      message: "Account created successfully. Use your Country + WhatsApp Number and Password to log in.",
      credentials: {
        country: registration.country,
        mobileNumber: registration.mobile
      },
      user: buildUserResponse({
        ...registration.toObject(),
        mobileNumber: registration.mobile
      })
    });
  } catch (error) {
    if (error?.code === 11000) {
      const duplicateField = Object.keys(error.keyPattern || {})[0];
      const duplicateMessage =
        duplicateField === "email"
          ? "An account with this Email Address already exists."
          : "An account with this Country + WhatsApp Number already exists.";
      return res.status(409).json({ message: duplicateMessage });
    }

    return res.status(500).json({
      message: error?.message || "Registration failed due to a server error."
    });
  }
}

app.post("/api/register", registerHandler);
app.post("/api/auth/register", registerHandler);

app.post("/api/newsletter", async (req, res) => {
  const { email } = req.body;

  if (!email || !email.includes("@")) {
    return res.status(400).json({ message: "A valid email is required." });
  }

  return res.status(201).json({ message: "Subscription received.", email });
});

function formatSystemStatusLabel(value = "") {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return "--";
  if (normalized === STATUS_PENDING || normalized === "pending") return "Pending Review";
  if (normalized === STATUS_ACTIVE) return "Active";
  if (normalized === STATUS_SUSPENDED || normalized === "inactive" || normalized === "deactive") return "Suspended";
  return normalized
    .split("-")
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

async function buildSuperAdminDashboardStats() {
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const [
    totalAdmins,
    totalAssociateMembers,
    totalOrders,
    pendingOrders,
    printingOrders,
    packagingOrders,
    dispatchOrders,
    completedOrders,
    pendingTopUpRequests,
    totalTopUpRequests,
    approvedTopUpAmountAggregate,
    walletDebitAggregate,
    recentUsers,
    recentOrders,
    recentTopUpRequests,
    recentWalletDebitOrders,
    pendingUsers,
    pendingOrderItems,
    pendingWalletRequests
  ] = await Promise.all([
    User.countDocuments({ role: ROLE_ADMIN }),
    User.countDocuments({ role: ROLE_ASSOCIATE_MEMBER }),
    Order.countDocuments({}),
    Order.countDocuments({ status: ORDER_STATUS_PENDING }),
    Order.countDocuments({ status: ORDER_STATUS_PRINTING }),
    Order.countDocuments({ status: ORDER_STATUS_PACKAGING }),
    Order.countDocuments({ status: ORDER_STATUS_DISPATCHED }),
    Order.countDocuments({ status: ORDER_STATUS_COMPLETED }),
    TopUpRequest.countDocuments({ status: "pending" }),
    TopUpRequest.countDocuments({}),
    TopUpRequest.aggregate([{ $match: { status: "approved" } }, { $group: { _id: null, total: { $sum: "$amount" } } }]),
    Order.aggregate([
      { $match: { walletDebitAmount: { $gt: 0 } } },
      { $group: { _id: null, totalAmount: { $sum: "$walletDebitAmount" }, totalCount: { $sum: 1 } } }
    ]),
    User.find({ role: { $in: [ROLE_ADMIN, ROLE_ASSOCIATE_MEMBER] } })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .select("ownerName role businessName mobile status updatedAt"),
    Order.find({})
      .sort({ orderedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("placedByUserId", "ownerName businessName")
      .select("orderNumber orderName customerName status orderedAt createdAt placedByUserId"),
    TopUpRequest.find({})
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("requestedByUserId", "ownerName businessName mobile role"),
    Order.find({ walletDebitAmount: { $gt: 0 } })
      .sort({ orderedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("placedByUserId", "ownerName businessName mobile role")
      .select("orderNumber orderName status orderedAt createdAt walletDebitAmount referenceNo placedByUserId"),
    User.find({ role: { $in: [ROLE_ADMIN, ROLE_ASSOCIATE_MEMBER] }, status: STATUS_PENDING })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .select("ownerName role businessName mobile status updatedAt"),
    Order.find({ status: ORDER_STATUS_PENDING })
      .sort({ updatedAt: -1, orderedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("placedByUserId", "ownerName businessName")
      .select("orderNumber orderName customerName status orderedAt updatedAt placedByUserId"),
    TopUpRequest.find({ status: "pending" })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("requestedByUserId", "ownerName businessName mobile role")
  ]);

  const totalUsers = totalAdmins + totalAssociateMembers + 1;
  const approvedTopUpTotal = Number(approvedTopUpAmountAggregate[0]?.total || 0);
  const walletDebitTotal = Number(walletDebitAggregate[0]?.totalAmount || 0);
  const walletDebitCount = Number(walletDebitAggregate[0]?.totalCount || 0);
  const totalWalletBalance = Math.max(0, Number((approvedTopUpTotal - walletDebitTotal).toFixed(2)));
  const totalWalletTransactions = totalTopUpRequests + walletDebitCount;
  const activeTodayUsers = await User.countDocuments({
    role: { $in: [ROLE_ADMIN, ROLE_ASSOCIATE_MEMBER] },
    $or: [{ lastLoginAt: { $gte: startOfToday } }, { lastActivityAt: { $gte: startOfToday } }]
  });

  const recentUserActivities = recentUsers.map((user) => ({
    id: user._id.toString(),
    name: user.ownerName || "--",
    role: normalizeRoleLabel(user.role || ""),
    businessName: user.businessName || "--",
    status: formatSystemStatusLabel(user.status),
    statusValue: String(user.status || "").trim().toLowerCase(),
    updatedOn: formatDateTime(user.updatedAt),
    detailsPath:
      user.role === ROLE_ADMIN
        ? `/dashboard/super-admin/user-management/admin-management/details/${user._id.toString()}`
        : `/dashboard/super-admin/user-management/associate-member-management/details/${user._id.toString()}`
  }));

  const recentOrderRows = recentOrders.map((order) => ({
    id: order._id.toString(),
    orderNumber: order.orderNumber ? `#${order.orderNumber}` : "--",
    orderName: order.orderName || "--",
    createdBy: order.placedByUserId?.ownerName || order.customerName || "--",
    status: normalizeOrderStatusLabel(order.status),
    statusValue: String(order.status || "").trim().toLowerCase(),
    orderedOn: formatDateTime(order.orderedAt || order.createdAt),
    detailsPath: `/dashboard/super-admin/order-management/details/${order._id.toString()}`
  }));

  const recentWalletCredits = recentTopUpRequests.map((request) => ({
    id: request._id.toString(),
    reference: `TPR-${request._id.toString().slice(-6).toUpperCase()}`,
    type: "Credit",
    memberName: request.requestedByUserId?.ownerName || request.requestedByDisplayName || "Super Admin",
    amount: Number(request.amount || 0).toFixed(2),
    status: normalizeTopUpStatusLabel(request.status),
    statusValue: String(request.status || "").trim().toLowerCase(),
    businessName: request.requestedByUserId?.businessName || "--",
    updatedOn: formatDateTime(request.updatedAt || request.createdAt),
    sortDate: request.updatedAt || request.createdAt || now,
    detailsPath: `/dashboard/super-admin/wallet-management/details/${request._id.toString()}`
  }));

  const recentWalletDebitsRows = recentWalletDebitOrders.map((order) => ({
    id: `wallet-debit-${order._id.toString()}`,
    reference: order.referenceNo || `ORD-${String(order.orderNumber || "").trim() || order._id.toString().slice(-6).toUpperCase()}`,
    type: "Debit",
    memberName: order.placedByUserId?.ownerName || order.customerName || "--",
    amount: Number(order.walletDebitAmount || 0).toFixed(2),
    status: normalizeOrderStatusLabel(order.status),
    statusValue: String(order.status || "").trim().toLowerCase(),
    businessName: order.placedByUserId?.businessName || "--",
    updatedOn: formatDateTime(order.orderedAt || order.createdAt),
    sortDate: order.orderedAt || order.createdAt || now,
    detailsPath: `/dashboard/super-admin/order-management/details/${order._id.toString()}`
  }));

  const recentWalletTransactions = [...recentWalletCredits, ...recentWalletDebitsRows]
    .sort((left, right) => new Date(right.sortDate || 0).getTime() - new Date(left.sortDate || 0).getTime())
    .slice(0, 6)
    .map(({ sortDate, ...item }) => item);

  const pendingTaskRows = [
    ...pendingUsers.map((user) => ({
      id: `user-${user._id.toString()}`,
      module: "User Management",
      title: `${user.ownerName || "--"} (${normalizeRoleLabel(user.role || "")})`,
      status: formatSystemStatusLabel(user.status),
      statusValue: String(user.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(user.updatedAt),
      actionLabel: "Review",
      routePath:
        user.role === ROLE_ADMIN
          ? `/dashboard/super-admin/user-management/admin-management/details/${user._id.toString()}`
          : `/dashboard/super-admin/user-management/associate-member-management/details/${user._id.toString()}`,
      sortDate: user.updatedAt || now
    })),
    ...pendingOrderItems.map((order) => ({
      id: `order-${order._id.toString()}`,
      module: "Order Management",
      title: order.orderNumber ? `Order #${order.orderNumber}` : order.orderName || "--",
      status: normalizeOrderStatusLabel(order.status),
      statusValue: String(order.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(order.updatedAt || order.orderedAt || order.createdAt),
      actionLabel: "Process",
      routePath: `/dashboard/super-admin/order-management/details/${order._id.toString()}`,
      sortDate: order.updatedAt || order.orderedAt || order.createdAt || now
    })),
    ...pendingWalletRequests.map((request) => ({
      id: `wallet-${request._id.toString()}`,
      module: "Wallet Management",
      title: `Top-up ${request.requestedByUserId?.ownerName || request.requestedByDisplayName || "Request"}`,
      status: normalizeTopUpStatusLabel(request.status),
      statusValue: String(request.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(request.updatedAt || request.createdAt),
      actionLabel: "Review",
      routePath: `/dashboard/super-admin/wallet-management/details/${request._id.toString()}`,
      sortDate: request.updatedAt || request.createdAt || now
    }))
  ]
    .sort((left, right) => new Date(right.sortDate || 0).getTime() - new Date(left.sortDate || 0).getTime())
    .slice(0, 10)
    .map(({ sortDate, ...item }) => item);

  return {
    overview: {
      totalUsers,
      totalAdmins,
      totalAssociateMembers,
      totalOrders,
      pendingOrders,
      printingOrders,
      packagingOrders,
      dispatchOrders,
      completedOrders,
      totalWalletBalance: totalWalletBalance.toFixed(2),
      totalWalletTransactions,
      pendingWalletRequests,
      activeTodayUsers
    },
    recentActivities: {
      users: recentUserActivities,
      orders: recentOrderRows,
      walletTransactions: recentWalletTransactions
    },
    pendingTasks: pendingTaskRows
  };
}

app.get("/api/super-admin/stats", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  return res.json(await buildSuperAdminDashboardStats());
});

app.get("/api/super-admin/admins", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchAdminList(req.query);
  return res.json(result);
});

app.get("/api/super-admin/admins/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "ownerName",
    "mobile",
    "email",
    "businessName",
    "address",
    "country",
    "state",
    "district",
    "city",
    "pinCode",
    "password",
    "status",
    "associateMemberAccessStatus"
  ].join(",");

  const sample = [
    "Admin User",
    "9876543210",
    "admin@example.com",
    "Sample Print House",
    "Plot 1 Main Road",
    "India",
    "Rajasthan",
    "Jaipur",
    "Jaipur",
    "302012",
    "Admin@123",
    "Active",
    "Enabled"
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="admin-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/admins/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = [
    "ownerName",
    "mobile",
    "email",
    "businessName",
    "address",
    "country",
    "state",
    "district",
    "city",
    "pinCode",
    "password"
  ];

  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  const seenEmail = new Set();
  const seenCountryMobile = new Set();

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvAdminRecord(record, rowNumber);
    const normalizedEmail = String(record.email || "").trim().toLowerCase();
    const normalizedCountry = String(record.country || "").trim();
    const normalizedMobile = String(record.mobile || "").replace(/\D/g, "");
    const uniqueCountryMobileKey = `${normalizedCountry}::${normalizedMobile}`;

    if (normalizedEmail && seenEmail.has(normalizedEmail)) {
      rowErrors.push(`Row ${rowNumber}: Duplicate email found inside CSV file.`);
    }
    if (normalizedCountry && normalizedMobile && seenCountryMobile.has(uniqueCountryMobileKey)) {
      rowErrors.push(`Row ${rowNumber}: Duplicate country and mobile number found inside CSV file.`);
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    seenEmail.add(normalizedEmail);
    seenCountryMobile.add(uniqueCountryMobileKey);
    preparedRecords.push({
      ownerName: String(record.ownerName).trim(),
      mobile: normalizedMobile,
      email: normalizedEmail,
      businessName: String(record.businessName).trim(),
      address: String(record.address).trim(),
      country: normalizedCountry,
      state: String(record.state).trim(),
      district: String(record.district).trim(),
      city: String(record.city).trim(),
      pinCode: String(record.pinCode).trim(),
      password: String(record.password).trim(),
      status: normalizeStatusInput(record.status) || STATUS_ACTIVE,
      associateMemberAccessEnabled: normalizeBoolean(record.associateMemberAccessStatus) ?? false
    });
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  const existingUsers = await User.find({
    $or: preparedRecords.flatMap((record) => [
      { email: record.email },
      { country: record.country, mobile: record.mobile }
    ])
  }).select("email country mobile");

  if (existingUsers.length > 0) {
    const existingMessages = existingUsers.map((user) => {
      if (preparedRecords.some((record) => record.email === user.email)) {
        return `Admin with email ${user.email} already exists.`;
      }
      return `Admin with mobile ${user.mobile} in ${user.country} already exists.`;
    });

    return res.status(409).json({
      message: "Some Admin records already exist.",
      errors: existingMessages
    });
  }

  const documents = [];
  for (const record of preparedRecords) {
    const passwordHash = await bcrypt.hash(record.password, 10);
    documents.push({
      ownerName: record.ownerName,
      mobile: record.mobile,
      email: record.email,
      businessName: record.businessName,
      address: record.address,
      country: record.country,
      state: record.state,
      district: record.district,
      city: record.city,
      pinCode: record.pinCode,
      passwordHash,
      referenceCode: "",
      gstNumber: "",
      services: [],
      termsAccepted: true,
      role: ROLE_ADMIN,
      status: record.status,
      associateMemberAccessEnabled: record.associateMemberAccessEnabled
    });
  }

  await User.insertMany(documents);

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Management",
    action: "Import Admin CSV",
    targetType: "Admin",
    targetId: "",
    targetName: `${documents.length} Admin records`,
    details: { createdCount: documents.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${documents.length} Admin record(s) created.`,
    createdCount: documents.length
  });
});

app.get("/api/super-admin/admins/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchAdminList({ ...req.query, page: 1, limit: 10000 });

  const worksheetData = result.items.map((item) => ({
    "Admin ID": item.adminId,
    "Admin Name": item.adminName,
    "Mobile Number": item.mobileNumber,
    "Admin Address": item.adminAddress,
    "Business/Firm Name": item.businessName,
    Status: item.status,
    "Associate Member Access Status": item.associateMemberAccessStatus,
    "Associate Member Count": item.associateMemberCount
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Admins");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="admin-management.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/admins/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchAdminList({ ...req.query, page: 1, limit: 10000 });

  const header = [
    "Admin ID",
    "Admin Name",
    "Mobile Number",
    "Admin Address",
    "Business/Firm Name",
    "Status",
    "Associate Member Access Status",
    "Associate Member Count"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.adminId,
        `"${item.adminName.replace(/"/g, '""')}"`,
        item.mobileNumber,
        `"${item.adminAddress.replace(/"/g, '""')}"`,
        `"${item.businessName.replace(/"/g, '""')}"`,
        item.status,
        item.associateMemberAccessStatus,
        item.associateMemberCount
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="admin-management.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/admins/options", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const admins = await User.find({ role: ROLE_ADMIN }).sort({ ownerName: 1 }).select("ownerName mobile status");
  return res.json({
    items: admins.map((admin) => ({
      id: admin._id.toString(),
      name: admin.ownerName,
      mobileNumber: admin.mobile,
      status: normalizeAdminStatusLabel(admin.status)
    }))
  });
});

app.post("/api/super-admin/admins", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const ownerName = String(req.body.ownerName || "").trim();
  const businessName = String(req.body.businessName || "").trim();
  const email = String(req.body.email || "").trim().toLowerCase();
  const country = String(req.body.country || "").trim();
  const state = String(req.body.state || "").trim();
  const district = String(req.body.district || "").trim();
  const city = String(req.body.city || "").trim();
  const pinCode = String(req.body.pinCode || "").trim();
  const address = String(req.body.address || "").trim();
  const password = String(req.body.password || "");
  const mobile = normalizeMobileNumber(country, req.body.mobile);
  const nextStatus = normalizeStatusInput(req.body.status) || STATUS_ACTIVE;
  const associateMemberAccessEnabled = normalizeBoolean(req.body.associateMemberAccessEnabled) ?? false;

  if (!ownerName || !businessName || !email || !country || !state || !district || !city || !pinCode || !address || !password) {
    return res.status(400).json({ message: "Complete Admin account details are required." });
  }

  if (!/^\S+@\S+\.\S+$/.test(email)) {
    return res.status(400).json({ message: "Enter a valid email address." });
  }

  if (!mobile || !/^\d{10,15}$/.test(mobile)) {
    return res.status(400).json({ message: "Enter a valid WhatsApp Number with 10 to 15 digits." });
  }

  if (password.length < 8) {
    return res.status(400).json({ message: "Password must be at least 8 characters long." });
  }

  const [existingEmailUser, existingMobileUser] = await Promise.all([
    User.findOne({ email }).select("_id"),
    User.findOne({ country, mobile }).select("_id")
  ]);

  if (existingEmailUser) {
    return res.status(409).json({ message: "An account with this Email Address already exists." });
  }

  if (existingMobileUser) {
    return res.status(409).json({ message: "An account with this Country + WhatsApp Number already exists." });
  }

  const passwordHash = await bcrypt.hash(password, 10);
  const admin = await User.create({
    ownerName,
    businessName,
    mobile,
    email,
    passwordHash,
    referenceCode: "",
    country,
    state,
    district,
    city,
    pinCode,
    gstNumber: "",
    address,
    services: [],
    termsAccepted: true,
    captchaVerified: true,
    role: ROLE_ADMIN,
    status: nextStatus,
    associateMemberAccessEnabled
  });

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Management",
    action: "Create Admin",
    targetType: "Admin",
    targetId: admin._id.toString(),
    targetName: admin.ownerName,
    details: {
      nextStatus: normalizeAdminStatusLabel(nextStatus),
      associateMemberAccessStatus: associateMemberAccessEnabled ? "Enabled" : "Disabled"
    }
  });

  return res.status(201).json({
    message: "Admin account created successfully.",
    admin: await buildAdminDetails(admin)
  });
});

app.get("/api/super-admin/admins/:adminId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.adminId)) {
    return res.status(400).json({ message: "Invalid Admin record ID." });
  }

  const admin = await User.findOne({ _id: req.params.adminId, role: ROLE_ADMIN });
  if (!admin) {
    return res.status(404).json({ message: "Admin record not found." });
  }

  return res.json({ admin: await buildAdminDetails(admin) });
});

app.get("/api/super-admin/admins/:adminId/associate-members", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.adminId)) {
    return res.status(400).json({ message: "Invalid Admin record ID." });
  }

  const admin = await User.findOne({ _id: req.params.adminId, role: ROLE_ADMIN });
  if (!admin) {
    return res.status(404).json({ message: "Admin record not found." });
  }

  const details = await buildAdminDetails(admin);
  return res.json({
    items: details.associateMembers
  });
});

app.delete("/api/super-admin/admins/:adminId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.adminId)) {
    return res.status(400).json({ message: "Invalid Admin record ID." });
  }

  const admin = await User.findOne({ _id: req.params.adminId, role: ROLE_ADMIN }).select("_id ownerName");
  if (!admin) {
    return res.status(404).json({ message: "Admin record not found." });
  }

  const assignedAssociateIds = await User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: admin._id }).distinct("_id");

  await Promise.all([
    User.updateMany(
      { role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: admin._id },
      { $set: { assignedAdminId: null, lastActivityAt: new Date() } }
    ),
    Order.updateMany(
      { assignedAdminId: admin._id },
      {
        $set: {
          assignedAdminId: null
        }
      }
    ),
    User.deleteOne({ _id: admin._id, role: ROLE_ADMIN })
  ]);

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Management",
    action: "Delete Admin",
    targetType: "Admin",
    targetId: admin._id.toString(),
    targetName: admin.ownerName,
    details: {
      unassignedAssociateMembers: assignedAssociateIds.length
    }
  });

  return res.json({
    message: "Admin deleted successfully. Assigned Associate Members are now managed directly under Super Admin.",
    unassignedAssociateMembers: assignedAssociateIds.length
  });
});

app.patch("/api/super-admin/admins/:adminId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.adminId)) {
    return res.status(400).json({ message: "Invalid Admin record ID." });
  }

  const nextStatus = normalizeStatusInput(req.body.status);
  if (!nextStatus) {
    return res.status(400).json({ message: "Status must be Active or Deactive." });
  }

  const admin = await User.findOneAndUpdate(
    { _id: req.params.adminId, role: ROLE_ADMIN },
    { $set: { status: nextStatus, lastActivityAt: new Date() } },
    { new: true }
  );

  if (!admin) {
    return res.status(404).json({ message: "Admin record not found." });
  }

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Management",
    action: nextStatus === STATUS_ACTIVE ? "Activate Admin" : "Deactivate Admin",
    targetType: "Admin",
    targetId: admin._id.toString(),
    targetName: admin.ownerName,
    details: {
      nextStatus: normalizeAdminStatusLabel(nextStatus)
    }
  });

  return res.json({
    message: `Admin account ${nextStatus === STATUS_ACTIVE ? "activated" : "deactivated"} successfully.`,
    admin: await buildAdminDetails(admin)
  });
});

app.patch("/api/super-admin/admins/:adminId/associate-access", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.adminId)) {
    return res.status(400).json({ message: "Invalid Admin record ID." });
  }

  const enabled = normalizeBoolean(req.body.enabled);
  if (enabled === null) {
    return res.status(400).json({ message: "Associate Member access must be Enabled or Disabled." });
  }

  const admin = await User.findOneAndUpdate(
    { _id: req.params.adminId, role: ROLE_ADMIN },
    { $set: { associateMemberAccessEnabled: enabled, lastActivityAt: new Date() } },
    { new: true }
  );

  if (!admin) {
    return res.status(404).json({ message: "Admin record not found." });
  }

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Management",
    action: enabled ? "Enable Associate Member Access" : "Disable Associate Member Access",
    targetType: "Admin",
    targetId: admin._id.toString(),
    targetName: admin.ownerName,
    details: {
      associateMemberAccessStatus: enabled ? "Enabled" : "Disabled"
    }
  });

  return res.json({
    message: `Associate Member access ${enabled ? "enabled" : "disabled"} successfully.`,
    admin: await buildAdminDetails(admin)
  });
});

app.get("/api/super-admin/admins/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const logs = await AuditLog.find({ module: "Admin Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get(
  "/api/super-admin/associate-members",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    const result = await fetchAssociateMemberList(req.query);
    // #region debug-point H3:associate-api-response
    fetch("http://127.0.0.1:7777/event",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({sessionId:"associate-table-api",runId:"pre",hypothesisId:"H3",location:"server.js:2318",msg:"[DEBUG] Associate Member API returning payload",data:{query:req.query,itemCount:Array.isArray(result?.items)?result.items.length:null,totalRecords:result?.pagination?.totalRecords??null,totalPages:result?.pagination?.totalPages??null,userRole:req.user?.role||null},ts:Date.now()})}).catch(()=>{});
    // #endregion
    return res.json(result);
  }
);

app.get(
  "/api/super-admin/associate-members/template.csv",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  (_req, res) => {
    const header = [
      "ownerName",
      "mobile",
      "email",
      "businessName",
      "address",
      "country",
      "state",
      "district",
      "city",
      "pinCode",
      "password",
      "status",
      "assignedAdminMobile"
    ].join(",");

    const sample = [
      "Associate Member",
      "9876500001",
      "associate@example.com",
      "Sample Bag House",
      "Plot 12 Main Road",
      "India",
      "Rajasthan",
      "Jaipur",
      "Jaipur",
      "302012",
      "Associate@123",
      "Active",
      "9876543210"
    ].join(",");

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", 'attachment; filename="associate-member-import-template.csv"');
    return res.send(`${header}\n${sample}`);
  }
);

app.post(
  "/api/super-admin/associate-members/import-csv",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    const { csvText } = req.body;

    if (!csvText || typeof csvText !== "string") {
      return res.status(400).json({ message: "CSV content is required." });
    }

    const rows = parseCsvText(csvText);
    if (rows.length < 2) {
      return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
    }

    const headers = rows[0];
    const requiredHeaders = [
      "ownerName",
      "mobile",
      "email",
      "businessName",
      "address",
      "country",
      "state",
      "district",
      "city",
      "pinCode",
      "password"
    ];

    const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
    if (missingHeaders.length > 0) {
      return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
    }

    const validationErrors = [];
    const preparedRecords = [];
    const seenEmail = new Set();
    const seenCountryMobile = new Set();

    for (let index = 1; index < rows.length; index += 1) {
      const rowValues = rows[index];
      const rowNumber = index + 1;
      const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

      const rowErrors = validateCsvAssociateMemberRecord(record, rowNumber);
      const normalizedEmail = String(record.email || "").trim().toLowerCase();
      const normalizedCountry = String(record.country || "").trim();
      const normalizedMobile = String(record.mobile || "").replace(/\D/g, "");
      const uniqueCountryMobileKey = `${normalizedCountry}::${normalizedMobile}`;

      if (normalizedEmail && seenEmail.has(normalizedEmail)) {
        rowErrors.push(`Row ${rowNumber}: Duplicate email found inside CSV file.`);
      }
      if (normalizedCountry && normalizedMobile && seenCountryMobile.has(uniqueCountryMobileKey)) {
        rowErrors.push(`Row ${rowNumber}: Duplicate country and mobile number found inside CSV file.`);
      }

      let assignedAdminId = null;
      if (record.assignedAdminMobile) {
        const normalizedAssignedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
        const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAssignedAdminMobile }).select("_id ownerName");
        if (!admin) {
          rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
        } else {
          assignedAdminId = admin._id;
        }
      }

      if (rowErrors.length > 0) {
        validationErrors.push(...rowErrors);
        continue;
      }

      seenEmail.add(normalizedEmail);
      seenCountryMobile.add(uniqueCountryMobileKey);
      preparedRecords.push({
        ownerName: String(record.ownerName).trim(),
        mobile: normalizedMobile,
        email: normalizedEmail,
        businessName: String(record.businessName).trim(),
        address: String(record.address).trim(),
        country: normalizedCountry,
        state: String(record.state).trim(),
        district: String(record.district).trim(),
        city: String(record.city).trim(),
        pinCode: String(record.pinCode).trim(),
        password: String(record.password).trim(),
        status: normalizeStatusInput(record.status) || STATUS_ACTIVE,
        assignedAdminId
      });
    }

    if (validationErrors.length > 0) {
      return res.status(400).json({
        message: "CSV validation failed.",
        errors: validationErrors
      });
    }

    const existingUsers = await User.find({
      $or: preparedRecords.flatMap((record) => [
        { email: record.email },
        { country: record.country, mobile: record.mobile }
      ])
    }).select("email country mobile");

    if (existingUsers.length > 0) {
      const existingMessages = existingUsers.map((user) => {
        if (preparedRecords.some((record) => record.email === user.email)) {
          return `Associate Member with email ${user.email} already exists.`;
        }
        return `Associate Member with mobile ${user.mobile} in ${user.country} already exists.`;
      });

      return res.status(409).json({
        message: "Some Associate Member records already exist.",
        errors: existingMessages
      });
    }

    const documents = [];
    for (const record of preparedRecords) {
      const passwordHash = await bcrypt.hash(record.password, 10);
      documents.push({
        ownerName: record.ownerName,
        mobile: record.mobile,
        email: record.email,
        businessName: record.businessName,
        address: record.address,
        country: record.country,
        state: record.state,
        district: record.district,
        city: record.city,
        pinCode: record.pinCode,
        passwordHash,
        referenceCode: "",
        gstNumber: "",
        services: ["Bag Printing"],
        termsAccepted: true,
        role: ROLE_ASSOCIATE_MEMBER,
        status: record.status,
        assignedAdminId: record.assignedAdminId
      });
    }

    await User.insertMany(documents);

    const actorName = await getActorName(req.auth);
    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Associate Member Management",
      action: "Import Associate Member CSV",
      targetType: "Associate Member",
      targetId: "",
      targetName: `${documents.length} Associate Member records`,
      details: { createdCount: documents.length }
    });

    return res.status(201).json({
      message: `CSV import completed successfully. ${documents.length} Associate Member record(s) created.`,
      createdCount: documents.length
    });
  }
);

app.get(
  "/api/super-admin/associate-members/export.xlsx",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    const result = await fetchAssociateMemberList({ ...req.query, page: 1, limit: 10000 });

    const worksheetData = result.items.map((item) => ({
      "Associate Member ID": item.associateMemberId,
      "Associate Member Name": item.associateMemberName,
      "Mobile Number": item.mobileNumber,
      "Assigned Admin Name": item.assignedAdminName,
      Status: item.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Associate Members");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="associate-member-management.xlsx"');
    return res.send(buffer);
  }
);

app.get(
  "/api/super-admin/associate-members/export.csv",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    const result = await fetchAssociateMemberList({ ...req.query, page: 1, limit: 10000 });

    const header = [
      "Associate Member ID",
      "Associate Member Name",
      "Mobile Number",
      "Assigned Admin Name",
      "Status"
    ];

    const csvRows = [
      header.join(","),
      ...result.items.map((item) =>
        [
          item.associateMemberId,
          `"${item.associateMemberName.replace(/"/g, '""')}"`,
          item.mobileNumber,
          `"${item.assignedAdminName.replace(/"/g, '""')}"`,
          item.status
        ].join(",")
      )
    ];

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", 'attachment; filename="associate-member-management.csv"');
    return res.send(csvRows.join("\n"));
  }
);

app.get(
  "/api/super-admin/associate-members/:associateMemberId",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
      return res.status(400).json({ message: "Invalid Associate Member record ID." });
    }

    const associateMember = await User.findOne({ _id: req.params.associateMemberId, role: ROLE_ASSOCIATE_MEMBER });
    if (!associateMember) {
      return res.status(404).json({ message: "Associate Member record not found." });
    }

    return res.json({ associateMember: await buildAssociateMemberDetails(associateMember) });
  }
);

app.patch(
  "/api/super-admin/associate-members/:associateMemberId/status",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
      return res.status(400).json({ message: "Invalid Associate Member record ID." });
    }

    const nextStatus = normalizeStatusInput(req.body.status);
    if (!nextStatus) {
      return res.status(400).json({ message: "Status must be Active or Deactive." });
    }

    const associateMember = await User.findOneAndUpdate(
      { _id: req.params.associateMemberId, role: ROLE_ASSOCIATE_MEMBER },
      { $set: { status: nextStatus, lastActivityAt: new Date() } },
      { new: true }
    );

    if (!associateMember) {
      return res.status(404).json({ message: "Associate Member record not found." });
    }

    const actorName = await getActorName(req.auth);
    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Associate Member Management",
      action: nextStatus === STATUS_ACTIVE ? "Activate Associate Member" : "Deactivate Associate Member",
      targetType: "Associate Member",
      targetId: associateMember._id.toString(),
      targetName: associateMember.ownerName,
      details: {
        nextStatus: normalizeAdminStatusLabel(nextStatus)
      }
    });

    return res.json({
      message: `Associate Member account ${nextStatus === STATUS_ACTIVE ? "activated" : "deactivated"} successfully.`,
      associateMember: await buildAssociateMemberDetails(associateMember)
    });
  }
);

app.patch(
  "/api/super-admin/associate-members/:associateMemberId/assign-admin",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
      return res.status(400).json({ message: "Invalid Associate Member record ID." });
    }

    const { adminId } = req.body;
    let nextAssignedAdminId = null;
    let assignedAdminName = "Not Assigned";

    if (adminId) {
      if (!mongoose.Types.ObjectId.isValid(adminId)) {
        return res.status(400).json({ message: "Invalid Admin record ID." });
      }

      const admin = await User.findOne({ _id: adminId, role: ROLE_ADMIN }).select("_id ownerName");
      if (!admin) {
        return res.status(404).json({ message: "Assigned Admin record not found." });
      }

      nextAssignedAdminId = admin._id;
      assignedAdminName = admin.ownerName;
    }

    const associateMember = await User.findOneAndUpdate(
      { _id: req.params.associateMemberId, role: ROLE_ASSOCIATE_MEMBER },
      { $set: { assignedAdminId: nextAssignedAdminId, lastActivityAt: new Date() } },
      { new: true }
    );

    if (!associateMember) {
      return res.status(404).json({ message: "Associate Member record not found." });
    }

    await syncAssociateMemberOrdersAdminAssignment(associateMember._id, nextAssignedAdminId);

    const actorName = await getActorName(req.auth);
    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Associate Member Management",
      action: nextAssignedAdminId ? "Assign Admin" : "Remove Assigned Admin",
      targetType: "Associate Member",
      targetId: associateMember._id.toString(),
      targetName: associateMember.ownerName,
      details: {
        assignedAdminId: nextAssignedAdminId?.toString?.() || "",
        assignedAdminName
      }
    });

    return res.json({
      message: nextAssignedAdminId
        ? `Assigned Admin updated to ${assignedAdminName}.`
        : "Assigned Admin removed successfully.",
      associateMember: await buildAssociateMemberDetails(associateMember)
    });
  }
);

app.patch(
  "/api/super-admin/associate-members/:associateMemberId/promote",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
      return res.status(400).json({ message: "Invalid Associate Member record ID." });
    }

    const associateMember = await User.findOneAndUpdate(
      { _id: req.params.associateMemberId, role: ROLE_ASSOCIATE_MEMBER },
      {
        $set: {
          role: ROLE_ADMIN,
          assignedAdminId: null,
          associateMemberAccessEnabled: false,
          lastActivityAt: new Date()
        }
      },
      { new: true }
    );

    if (!associateMember) {
      return res.status(404).json({ message: "Associate Member record not found." });
    }

    await syncAssociateMemberOrdersAdminAssignment(associateMember._id, null);

    const actorName = await getActorName(req.auth);
    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Associate Member Management",
      action: "Promote Associate Member To Admin",
      targetType: "Associate Member",
      targetId: associateMember._id.toString(),
      targetName: associateMember.ownerName,
      details: {
        nextRole: ROLE_ADMIN
      }
    });

    return res.json({
      message: "Associate Member promoted to Admin successfully."
    });
  }
);

app.patch(
  "/api/super-admin/associate-members/:associateMemberId/revert-role",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
      return res.status(400).json({ message: "Invalid user record ID." });
    }

    const { adminId } = req.body;
    let nextAssignedAdminId = null;

    if (adminId) {
      if (!mongoose.Types.ObjectId.isValid(adminId)) {
        return res.status(400).json({ message: "Invalid Admin record ID." });
      }

      const admin = await User.findOne({ _id: adminId, role: ROLE_ADMIN }).select("_id");
      if (!admin) {
        return res.status(404).json({ message: "Assigned Admin record not found." });
      }
      nextAssignedAdminId = admin._id;
    }

    const user = await User.findOneAndUpdate(
      { _id: req.params.associateMemberId, role: ROLE_ADMIN },
      {
        $set: {
          role: ROLE_ASSOCIATE_MEMBER,
          assignedAdminId: nextAssignedAdminId,
          associateMemberAccessEnabled: false,
          lastActivityAt: new Date()
        }
      },
      { new: true }
    );

    if (!user) {
      return res.status(404).json({ message: "Admin record not found for role revert." });
    }

    await syncAssociateMemberOrdersAdminAssignment(user._id, nextAssignedAdminId);

    const actorName = await getActorName(req.auth);
    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: "Associate Member Management",
      action: "Revoke Admin Privileges",
      targetType: "Associate Member",
      targetId: user._id.toString(),
      targetName: user.ownerName,
      details: {
        nextRole: ROLE_ASSOCIATE_MEMBER,
        assignedAdminId: nextAssignedAdminId?.toString?.() || ""
      }
    });

    return res.json({
      message: "Admin privileges revoked and user reverted to Associate Member successfully."
    });
  }
);

app.get(
  "/api/super-admin/associate-members/activity/logs",
  authenticate,
  requireRole(ROLE_SUPER_ADMIN),
  ensureDatabaseConnected,
  async (_req, res) => {
    const logs = await AuditLog.find({ module: "Associate Member Management" }).sort({ createdAt: -1 }).limit(100);
    return res.json({
      items: logs.map((log) => ({
        id: log._id.toString(),
        action: log.action,
        actorName: log.actorName,
        targetName: log.targetName,
        createdAt: log.createdAt,
        details: log.details
      }))
    });
  }
);

app.get("/api/super-admin/orders", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList(req.query);
  return res.json(result);
});

app.get("/api/super-admin/orders/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "orderName",
    "customerName",
    "customerMobile",
    "orderDetailsOverview",
    "status",
    "designSubmissionSource",
    "designFileName",
    "designFileUrl",
    "referenceNo",
    "orderDateTime",
    "assignedAssociateMemberMobile",
    "assignedAdminMobile"
  ].join(",");

  const sample = [
    "Sample Bag Printing Order",
    "Ravi Sharma",
    "9876500001",
    "500 cotton bags, 2 color print, front and back side artwork",
    "Printing",
    "Online Upload",
    "bag-design.pdf",
    "https://example.com/designs/bag-design.pdf",
    "ORD-REF-1001",
    "2026-06-18T10:30:00Z",
    "9876500002",
    "9876500003"
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="orders-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/orders/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  let nextOrderNumber = await getNextOrderNumber();

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvOrderRecord(record, rowNumber);

    let assignedAssociateMemberId = null;
    let assignedAdminId = null;
    let placedByUserId = null;

    const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
    if (normalizedMobile) {
      const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
      if (placedByUser) {
        placedByUserId = placedByUser._id;
      }
    }

    if (record.assignedAssociateMemberMobile) {
      const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
      const associateMember = await User.findOne({
        role: ROLE_ASSOCIATE_MEMBER,
        mobile: normalizedAssociateMobile
      }).select("_id assignedAdminId");

      if (!associateMember) {
        rowErrors.push(
          `Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`
        );
      } else {
        assignedAssociateMemberId = associateMember._id;
        if (associateMember.assignedAdminId) {
          assignedAdminId = associateMember.assignedAdminId;
        }
      }
    }

    if (record.assignedAdminMobile) {
      const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
      const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
      if (!admin) {
        rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
      } else {
        assignedAdminId = admin._id;
      }
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    const normalizedStatus = normalizeOrderStatusInput(record.status) || ORDER_STATUS_PENDING;
    const normalizedDesignSource =
      normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;

    preparedRecords.push({
      orderNumber: nextOrderNumber,
      orderName: String(record.orderName).trim(),
      customerName: String(record.customerName).trim(),
      customerMobile: normalizedMobile,
      orderDetailsOverview: String(record.orderDetailsOverview).trim(),
      orderedAt: record.orderDateTime ? new Date(record.orderDateTime) : new Date(),
      status: normalizedStatus,
      designSubmissionSource: normalizedDesignSource,
      designFileName: String(record.designFileName || "").trim(),
      designFileUrl: String(record.designFileUrl || "").trim(),
      referenceNo: String(record.referenceNo || "").trim(),
      placedByUserId,
      assignedAssociateMemberId,
      assignedAdminId,
      statusHistory: [
        {
          status: normalizedStatus,
          note: "Order imported from CSV.",
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: await getActorName(req.auth)
        }
      ]
    });

    nextOrderNumber += 5;
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  if (preparedRecords.length === 0) {
    return res.status(400).json({ message: "No valid order records were found in the CSV file." });
  }

  await Order.insertMany(preparedRecords);

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "All Orders Management",
    action: "Import Orders CSV",
    targetType: "Order",
    targetId: "",
    targetName: `${preparedRecords.length} Order records`,
    details: { createdCount: preparedRecords.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${preparedRecords.length} Order record(s) created.`,
    createdCount: preparedRecords.length
  });
});

app.get("/api/super-admin/orders/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000 });

  const worksheetData = result.items.map((item) => ({
    "Order Number": item.orderNumber,
    "Order Name": item.orderName,
    "Order Date & Time": item.orderDateTime,
    "Order Details Overview": item.orderDetailsOverview,
    "Customer Name": item.customerName,
    "Mobile Number": item.mobileNumber,
    "Current Status": item.currentStatus,
    "Design File Source": item.designFileSource
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "All Orders");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="all-orders.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/orders/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000 });

  const header = [
    "Order Number",
    "Order Name",
    "Order Date & Time",
    "Order Details Overview",
    "Customer Name",
    "Mobile Number",
    "Current Status",
    "Design File Source"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.orderNumber,
        `"${String(item.orderName || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
        `"${String(item.customerName || "").replace(/"/g, '""')}"`,
        item.mobileNumber,
        item.currentStatus,
        item.designFileSource
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="all-orders.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/orders/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const logs = await AuditLog.find({ module: "All Orders Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get("/api/super-admin/orders/pending", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, allowedStatuses: PENDING_ORDER_ROUTE_STATUSES });
  return res.json(result);
});

app.get("/api/super-admin/orders/pending/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "orderName",
    "customerName",
    "customerMobile",
    "orderDetailsOverview",
    "status",
    "designSubmissionSource",
    "designFileName",
    "designFileUrl",
    "referenceNo",
    "orderDateTime",
    "assignedAssociateMemberMobile",
    "assignedAdminMobile"
  ].join(",");

  const sample = [
    "Pending Bag Printing Order",
    "Vikas Sharma",
    "9876500011",
    "250 shopping bags with front-side logo printing",
    "Printing",
    "Online Upload",
    "pending-design.pdf",
    "https://example.com/designs/pending-design.pdf",
    "ORD-PND-1001",
    "2026-06-18T11:30:00Z",
    "9876500002",
    "9876500003"
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="pending-orders-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/orders/pending/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  let nextOrderNumber = await getNextOrderNumber();
  const actorName = await getActorName(req.auth);

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvOrderRecord(record, rowNumber, PENDING_ORDER_ROUTE_STATUSES);

    let assignedAssociateMemberId = null;
    let assignedAdminId = null;
    let placedByUserId = null;

    const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
    if (normalizedMobile) {
      const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
      if (placedByUser) {
        placedByUserId = placedByUser._id;
      }
    }

    if (record.assignedAssociateMemberMobile) {
      const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
      const associateMember = await User.findOne({
        role: ROLE_ASSOCIATE_MEMBER,
        mobile: normalizedAssociateMobile
      }).select("_id assignedAdminId");

      if (!associateMember) {
        rowErrors.push(
          `Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`
        );
      } else {
        assignedAssociateMemberId = associateMember._id;
        if (associateMember.assignedAdminId) {
          assignedAdminId = associateMember.assignedAdminId;
        }
      }
    }

    if (record.assignedAdminMobile) {
      const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
      const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
      if (!admin) {
        rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
      } else {
        assignedAdminId = admin._id;
      }
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    const normalizedStatus = normalizeOrderStatusInput(record.status) || ORDER_STATUS_PRINTING;
    const normalizedDesignSource =
      normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;

    preparedRecords.push({
      orderNumber: nextOrderNumber,
      orderName: String(record.orderName).trim(),
      customerName: String(record.customerName).trim(),
      customerMobile: normalizedMobile,
      orderDetailsOverview: String(record.orderDetailsOverview).trim(),
      orderedAt: record.orderDateTime ? new Date(record.orderDateTime) : new Date(),
      status: normalizedStatus,
      designSubmissionSource: normalizedDesignSource,
      designFileName: String(record.designFileName || "").trim(),
      designFileUrl: String(record.designFileUrl || "").trim(),
      referenceNo: String(record.referenceNo || "").trim(),
      placedByUserId,
      assignedAssociateMemberId,
      assignedAdminId,
      statusHistory: [
        {
          status: normalizedStatus,
          note: "Pending Order imported from CSV.",
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      ]
    });

    nextOrderNumber += 5;
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  if (preparedRecords.length === 0) {
    return res.status(400).json({ message: "No valid pending order records were found in the CSV file." });
  }

  await Order.insertMany(preparedRecords);

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Pending Orders Management",
    action: "Import Pending Orders CSV",
    targetType: "Order",
    targetId: "",
    targetName: `${preparedRecords.length} Pending Order records`,
    details: { createdCount: preparedRecords.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${preparedRecords.length} Pending Order record(s) created.`,
    createdCount: preparedRecords.length
  });
});

app.get("/api/super-admin/orders/pending/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PENDING_ORDER_ROUTE_STATUSES });

  const worksheetData = result.items.map((item) => ({
    "Order Number": item.orderNumber,
    "Order Name": item.orderName,
    "Order Date & Time": item.orderDateTime,
    "Order Details Overview": item.orderDetailsOverview,
    "Customer Name": item.customerName,
    "Mobile Number": item.mobileNumber,
    "Current Status": item.currentStatus,
    "Design File Source": item.designFileSource
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Pending Orders");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="pending-orders.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/orders/pending/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PENDING_ORDER_ROUTE_STATUSES });

  const header = [
    "Order Number",
    "Order Name",
    "Order Date & Time",
    "Order Details Overview",
    "Customer Name",
    "Mobile Number",
    "Current Status",
    "Design File Source"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.orderNumber,
        `"${String(item.orderName || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
        `"${String(item.customerName || "").replace(/"/g, '""')}"`,
        item.mobileNumber,
        item.currentStatus,
        item.designFileSource
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="pending-orders.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/orders/pending/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const logs = await AuditLog.find({ module: "Pending Orders Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get("/api/super-admin/orders/pending/:orderId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Pending Order record ID." });
  }

  await ensureOrderNumbersForExistingOrders();

  const order = await Order.findOne({ _id: req.params.orderId, status: { $in: PENDING_ORDER_ROUTE_STATUSES } });
  if (!order) {
    return res.status(404).json({ message: "Pending Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.patch("/api/super-admin/orders/pending/:orderId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Pending Order record ID." });
  }

  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (!PENDING_ORDER_ROUTE_STATUSES.includes(nextStatus)) {
    return res.status(400).json({
      message: "Pending Order status must be Printing or Packaging."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();

  const order = await Order.findOneAndUpdate(
    { _id: req.params.orderId, status: { $in: PENDING_ORDER_ROUTE_STATUSES } },
    {
      $set: {
        status: nextStatus,
        updatedAt: new Date()
      },
      $push: {
        statusHistory: {
          status: nextStatus,
          note,
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      }
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Pending Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Pending Orders Management",
    action: `Update Pending Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      note
    }
  });

  return res.json({
    message: `Pending Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

app.get("/api/super-admin/orders/printing", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, allowedStatuses: PRINTING_ORDER_ROUTE_STATUSES });
  return res.json(result);
});

app.get("/api/super-admin/orders/printing/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "orderName",
    "customerName",
    "customerMobile",
    "orderDetailsOverview",
    "status",
    "designSubmissionSource",
    "designFileName",
    "designFileUrl",
    "referenceNo",
    "orderDateTime",
    "assignedAssociateMemberMobile",
    "assignedAdminMobile"
  ].join(",");

  const sample = [
    "Printing Stage Order",
    "Neha Gupta",
    "9876500021",
    "300 printed carry bags with one-color branding",
    "Printing",
    "Online Upload",
    "printing-design.pdf",
    "https://example.com/designs/printing-design.pdf",
    "ORD-PRT-1001",
    "2026-06-18T12:30:00Z",
    "9876500002",
    "9876500003"
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="printing-orders-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/orders/printing/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  let nextOrderNumber = await getNextOrderNumber();
  const actorName = await getActorName(req.auth);

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvOrderRecord(record, rowNumber, PRINTING_ORDER_ROUTE_STATUSES);

    let assignedAssociateMemberId = null;
    let assignedAdminId = null;
    let placedByUserId = null;

    const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
    if (normalizedMobile) {
      const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
      if (placedByUser) {
        placedByUserId = placedByUser._id;
      }
    }

    if (record.assignedAssociateMemberMobile) {
      const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
      const associateMember = await User.findOne({
        role: ROLE_ASSOCIATE_MEMBER,
        mobile: normalizedAssociateMobile
      }).select("_id assignedAdminId");

      if (!associateMember) {
        rowErrors.push(
          `Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`
        );
      } else {
        assignedAssociateMemberId = associateMember._id;
        if (associateMember.assignedAdminId) {
          assignedAdminId = associateMember.assignedAdminId;
        }
      }
    }

    if (record.assignedAdminMobile) {
      const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
      const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
      if (!admin) {
        rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
      } else {
        assignedAdminId = admin._id;
      }
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    const normalizedStatus = normalizeOrderStatusInput(record.status) || ORDER_STATUS_PRINTING;
    const normalizedDesignSource =
      normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;

    preparedRecords.push({
      orderNumber: nextOrderNumber,
      orderName: String(record.orderName).trim(),
      customerName: String(record.customerName).trim(),
      customerMobile: normalizedMobile,
      orderDetailsOverview: String(record.orderDetailsOverview).trim(),
      orderedAt: record.orderDateTime ? new Date(record.orderDateTime) : new Date(),
      status: normalizedStatus,
      designSubmissionSource: normalizedDesignSource,
      designFileName: String(record.designFileName || "").trim(),
      designFileUrl: String(record.designFileUrl || "").trim(),
      referenceNo: String(record.referenceNo || "").trim(),
      placedByUserId,
      assignedAssociateMemberId,
      assignedAdminId,
      statusHistory: [
        {
          status: normalizedStatus,
          note: "Printing Order imported from CSV.",
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      ]
    });

    nextOrderNumber += 5;
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  if (preparedRecords.length === 0) {
    return res.status(400).json({ message: "No valid Printing Order records were found in the CSV file." });
  }

  await Order.insertMany(preparedRecords);

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Printing Orders Management",
    action: "Import Printing Orders CSV",
    targetType: "Order",
    targetId: "",
    targetName: `${preparedRecords.length} Printing Order records`,
    details: { createdCount: preparedRecords.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${preparedRecords.length} Printing Order record(s) created.`,
    createdCount: preparedRecords.length
  });
});

app.get("/api/super-admin/orders/printing/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PRINTING_ORDER_ROUTE_STATUSES });

  const worksheetData = result.items.map((item) => ({
    "Order Number": item.orderNumber,
    "Order Name": item.orderName,
    "Order Date & Time": item.orderDateTime,
    "Order Details Overview": item.orderDetailsOverview,
    "Customer Name": item.customerName,
    "Mobile Number": item.mobileNumber,
    "Current Status": item.currentStatus,
    "Design File Source": item.designFileSource
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Printing Orders");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="printing-orders.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/orders/printing/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PRINTING_ORDER_ROUTE_STATUSES });

  const header = [
    "Order Number",
    "Order Name",
    "Order Date & Time",
    "Order Details Overview",
    "Customer Name",
    "Mobile Number",
    "Current Status",
    "Design File Source"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.orderNumber,
        `"${String(item.orderName || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
        `"${String(item.customerName || "").replace(/"/g, '""')}"`,
        item.mobileNumber,
        item.currentStatus,
        item.designFileSource
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="printing-orders.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/orders/printing/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const logs = await AuditLog.find({ module: "Printing Orders Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get("/api/super-admin/orders/printing/:orderId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Printing Order record ID." });
  }

  await ensureOrderNumbersForExistingOrders();

  const order = await Order.findOne({ _id: req.params.orderId, status: { $in: PRINTING_ORDER_ROUTE_STATUSES } });
  if (!order) {
    return res.status(404).json({ message: "Printing Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.patch("/api/super-admin/orders/printing/:orderId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Printing Order record ID." });
  }

  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (nextStatus !== ORDER_STATUS_PACKAGING) {
    return res.status(400).json({
      message: "Printing Order status can only be moved to Packaging."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();

  const order = await Order.findOneAndUpdate(
    { _id: req.params.orderId, status: { $in: PRINTING_ORDER_ROUTE_STATUSES } },
    {
      $set: {
        status: nextStatus,
        updatedAt: new Date()
      },
      $push: {
        statusHistory: {
          status: nextStatus,
          note,
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      }
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Printing Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Printing Orders Management",
    action: `Update Printing Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      note
    }
  });

  return res.json({
    message: `Printing Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

app.get("/api/super-admin/orders/packaging", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, allowedStatuses: PACKAGING_ORDER_ROUTE_STATUSES });
  return res.json(result);
});

app.get("/api/super-admin/orders/packaging/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "orderName",
    "customerName",
    "customerMobile",
    "orderDetailsOverview",
    "status",
    "designSubmissionSource",
    "designFileName",
    "designFileUrl",
    "referenceNo",
    "orderDateTime",
    "assignedAssociateMemberMobile",
    "assignedAdminMobile"
  ].join(",");

  const sample = [
    "Packaging Stage Order",
    "Amit Verma",
    "9876500031",
    "400 laminated bags ready for final packaging and dispatch",
    "Packaging",
    "Online Upload",
    "packaging-design.pdf",
    "https://example.com/designs/packaging-design.pdf",
    "ORD-PKG-1001",
    "2026-06-18T13:30:00Z",
    "9876500002",
    "9876500003"
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="packaging-orders-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/orders/packaging/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  let nextOrderNumber = await getNextOrderNumber();
  const actorName = await getActorName(req.auth);

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvOrderRecord(record, rowNumber, PACKAGING_ORDER_ROUTE_STATUSES);

    let assignedAssociateMemberId = null;
    let assignedAdminId = null;
    let placedByUserId = null;

    const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
    if (normalizedMobile) {
      const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
      if (placedByUser) {
        placedByUserId = placedByUser._id;
      }
    }

    if (record.assignedAssociateMemberMobile) {
      const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
      const associateMember = await User.findOne({
        role: ROLE_ASSOCIATE_MEMBER,
        mobile: normalizedAssociateMobile
      }).select("_id assignedAdminId");

      if (!associateMember) {
        rowErrors.push(
          `Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`
        );
      } else {
        assignedAssociateMemberId = associateMember._id;
        if (associateMember.assignedAdminId) {
          assignedAdminId = associateMember.assignedAdminId;
        }
      }
    }

    if (record.assignedAdminMobile) {
      const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
      const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
      if (!admin) {
        rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
      } else {
        assignedAdminId = admin._id;
      }
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    const normalizedStatus = normalizeOrderStatusInput(record.status) || ORDER_STATUS_PACKAGING;
    const normalizedDesignSource =
      normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;

    preparedRecords.push({
      orderNumber: nextOrderNumber,
      orderName: String(record.orderName).trim(),
      customerName: String(record.customerName).trim(),
      customerMobile: normalizedMobile,
      orderDetailsOverview: String(record.orderDetailsOverview).trim(),
      orderedAt: record.orderDateTime ? new Date(record.orderDateTime) : new Date(),
      status: normalizedStatus,
      designSubmissionSource: normalizedDesignSource,
      designFileName: String(record.designFileName || "").trim(),
      designFileUrl: String(record.designFileUrl || "").trim(),
      referenceNo: String(record.referenceNo || "").trim(),
      placedByUserId,
      assignedAssociateMemberId,
      assignedAdminId,
      statusHistory: [
        {
          status: normalizedStatus,
          note: "Packaging Order imported from CSV.",
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      ]
    });

    nextOrderNumber += 5;
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  if (preparedRecords.length === 0) {
    return res.status(400).json({ message: "No valid Packaging Order records were found in the CSV file." });
  }

  await Order.insertMany(preparedRecords);

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Packaging Orders Management",
    action: "Import Packaging Orders CSV",
    targetType: "Order",
    targetId: "",
    targetName: `${preparedRecords.length} Packaging Order records`,
    details: { createdCount: preparedRecords.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${preparedRecords.length} Packaging Order record(s) created.`,
    createdCount: preparedRecords.length
  });
});

app.get("/api/super-admin/orders/packaging/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PACKAGING_ORDER_ROUTE_STATUSES });

  const worksheetData = result.items.map((item) => ({
    "Order Number": item.orderNumber,
    "Order Name": item.orderName,
    "Order Date & Time": item.orderDateTime,
    "Order Details Overview": item.orderDetailsOverview,
    "Customer Name": item.customerName,
    "Mobile Number": item.mobileNumber,
    "Current Status": item.currentStatus,
    "Design File Source": item.designFileSource
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Packaging Orders");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="packaging-orders.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/orders/packaging/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: PACKAGING_ORDER_ROUTE_STATUSES });

  const header = [
    "Order Number",
    "Order Name",
    "Order Date & Time",
    "Order Details Overview",
    "Customer Name",
    "Mobile Number",
    "Current Status",
    "Design File Source"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.orderNumber,
        `"${String(item.orderName || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
        `"${String(item.customerName || "").replace(/"/g, '""')}"`,
        item.mobileNumber,
        item.currentStatus,
        item.designFileSource
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="packaging-orders.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/orders/packaging/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const logs = await AuditLog.find({ module: "Packaging Orders Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get("/api/super-admin/orders/packaging/:orderId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Packaging Order record ID." });
  }

  await ensureOrderNumbersForExistingOrders();

  const order = await Order.findOne({ _id: req.params.orderId, status: { $in: PACKAGING_ORDER_ROUTE_STATUSES } });
  if (!order) {
    return res.status(404).json({ message: "Packaging Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.patch("/api/super-admin/orders/packaging/:orderId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Packaging Order record ID." });
  }

  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (nextStatus !== ORDER_STATUS_DISPATCHED) {
    return res.status(400).json({
      message: "Packaging Order status can only be moved to Dispatched."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();

  const order = await Order.findOneAndUpdate(
    { _id: req.params.orderId, status: { $in: PACKAGING_ORDER_ROUTE_STATUSES } },
    {
      $set: {
        status: nextStatus,
        dispatchDateTime: req.body.dispatchDateTime ? new Date(req.body.dispatchDateTime) : new Date(),
        courierName: String(req.body.courierName || "").trim(),
        courierTrackingNumber: String(req.body.courierTrackingNumber || "").trim(),
        courierTrackingUrl: String(req.body.courierTrackingUrl || "").trim(),
        dispatchNotes: String(req.body.dispatchNotes || note || "Order dispatched from packaging stage.").trim(),
        deliveryStatus: String(req.body.deliveryStatus || "dispatched").trim().toLowerCase(),
        updatedAt: new Date()
      },
      $push: {
        statusHistory: {
          status: nextStatus,
          note,
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        },
        dispatchHistory: {
          status: nextStatus,
          note: String(req.body.dispatchNotes || note || "Order dispatched from packaging stage.").trim(),
          courierName: String(req.body.courierName || "").trim(),
          courierTrackingNumber: String(req.body.courierTrackingNumber || "").trim(),
          courierTrackingUrl: String(req.body.courierTrackingUrl || "").trim(),
          deliveryStatus: String(req.body.deliveryStatus || "dispatched").trim().toLowerCase(),
          eventAt: req.body.dispatchDateTime ? new Date(req.body.dispatchDateTime) : new Date(),
          updatedByRole: req.auth.role,
          updatedById: req.auth.sub,
          updatedByName: actorName
        }
      }
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Packaging Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Packaging Orders Management",
    action: `Update Packaging Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      note,
      courierName: order.courierName || "",
      courierTrackingNumber: order.courierTrackingNumber || "",
      dispatchDateTime: order.dispatchDateTime
    }
  });

  return res.json({
    message: `Packaging Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

app.get("/api/super-admin/orders/dispatched", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, allowedStatuses: DISPATCHED_ORDER_ROUTE_STATUSES });
  return res.json(result);
});

app.get("/api/super-admin/orders/dispatched/template.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
  const header = [
    "orderName",
    "customerName",
    "customerMobile",
    "orderDetailsOverview",
    "status",
    "designSubmissionSource",
    "designFileName",
    "designFileUrl",
    "referenceNo",
    "orderDateTime",
    "assignedAssociateMemberMobile",
    "assignedAdminMobile",
    "courierName",
    "courierTrackingNumber",
    "courierTrackingUrl",
    "dispatchDateTime",
    "dispatchNotes",
    "deliveryStatus",
    "deliveryDateTime"
  ].join(",");

  const sample = [
    "Dispatched Festival Bags",
    "Mohan Traders",
    "9876500041",
    "600 festival bags handed over to courier partner for delivery",
    "Dispatched",
    "Online Upload",
    "dispatch-design.pdf",
    "https://example.com/designs/dispatch-design.pdf",
    "ORD-DSP-1001",
    "2026-06-18T15:30:00Z",
    "9876500002",
    "9876500003",
    "Blue Dart",
    "BD123456789",
    "https://example.com/track/BD123456789",
    "2026-06-19T10:00:00Z",
    "Shipment picked up from Jaipur hub",
    "In Transit",
    ""
  ].join(",");

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="dispatched-orders-import-template.csv"');
  return res.send(`${header}\n${sample}`);
});

app.post("/api/super-admin/orders/dispatched/import-csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const { csvText } = req.body;

  if (!csvText || typeof csvText !== "string") {
    return res.status(400).json({ message: "CSV content is required." });
  }

  const rows = parseCsvText(csvText);
  if (rows.length < 2) {
    return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
  }

  const headers = rows[0];
  const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
  }

  const validationErrors = [];
  const preparedRecords = [];
  let nextOrderNumber = await getNextOrderNumber();
  const actorName = await getActorName(req.auth);

  for (let index = 1; index < rows.length; index += 1) {
    const rowValues = rows[index];
    const rowNumber = index + 1;
    const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

    const rowErrors = validateCsvOrderRecord(record, rowNumber, DISPATCHED_ORDER_ROUTE_STATUSES);
    let assignedAssociateMemberId = null;
    let assignedAdminId = null;
    let placedByUserId = null;

    const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
    if (normalizedMobile) {
      const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
      if (placedByUser) {
        placedByUserId = placedByUser._id;
      }
    }

    if (record.assignedAssociateMemberMobile) {
      const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
      const associateMember = await User.findOne({
        role: ROLE_ASSOCIATE_MEMBER,
        mobile: normalizedAssociateMobile
      }).select("_id assignedAdminId");

      if (!associateMember) {
        rowErrors.push(`Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`);
      } else {
        assignedAssociateMemberId = associateMember._id;
        if (associateMember.assignedAdminId) {
          assignedAdminId = associateMember.assignedAdminId;
        }
      }
    }

    if (record.assignedAdminMobile) {
      const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
      const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
      if (!admin) {
        rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
      } else {
        assignedAdminId = admin._id;
      }
    }

    if (rowErrors.length > 0) {
      validationErrors.push(...rowErrors);
      continue;
    }

    const normalizedDesignSource =
      normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;
    const normalizedDeliveryStatus = String(record.deliveryStatus || "dispatched").trim().toLowerCase();
    const dispatchDateTime = record.dispatchDateTime ? new Date(record.dispatchDateTime) : new Date();

    preparedRecords.push({
      orderNumber: nextOrderNumber,
      orderName: String(record.orderName).trim(),
      customerName: String(record.customerName).trim(),
      customerMobile: normalizedMobile,
      orderDetailsOverview: String(record.orderDetailsOverview).trim(),
      orderedAt: record.orderDateTime ? new Date(record.orderDateTime) : new Date(),
      status: ORDER_STATUS_DISPATCHED,
      designSubmissionSource: normalizedDesignSource,
      designFileName: String(record.designFileName || "").trim(),
      designFileUrl: String(record.designFileUrl || "").trim(),
      referenceNo: String(record.referenceNo || "").trim(),
      placedByUserId,
      assignedAssociateMemberId,
      assignedAdminId,
      courierName: String(record.courierName || "").trim(),
      courierTrackingNumber: String(record.courierTrackingNumber || "").trim(),
      courierTrackingUrl: String(record.courierTrackingUrl || "").trim(),
      dispatchDateTime,
      dispatchNotes: String(record.dispatchNotes || "Dispatched Order imported from CSV.").trim(),
      deliveryStatus: normalizedDeliveryStatus,
      deliveryDateTime: record.deliveryDateTime ? new Date(record.deliveryDateTime) : null,
      statusHistory: [
        {
          status: ORDER_STATUS_DISPATCHED,
          note: String(record.dispatchNotes || "Dispatched Order imported from CSV.").trim(),
          changedAt: dispatchDateTime,
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      ],
      dispatchHistory: [
        {
          status: ORDER_STATUS_DISPATCHED,
          note: String(record.dispatchNotes || "Dispatched Order imported from CSV.").trim(),
          courierName: String(record.courierName || "").trim(),
          courierTrackingNumber: String(record.courierTrackingNumber || "").trim(),
          courierTrackingUrl: String(record.courierTrackingUrl || "").trim(),
          deliveryStatus: normalizedDeliveryStatus,
          eventAt: dispatchDateTime,
          updatedByRole: req.auth.role,
          updatedById: req.auth.sub,
          updatedByName: actorName
        }
      ]
    });

    nextOrderNumber += 5;
  }

  if (validationErrors.length > 0) {
    return res.status(400).json({
      message: "CSV validation failed.",
      errors: validationErrors
    });
  }

  if (preparedRecords.length === 0) {
    return res.status(400).json({ message: "No valid Dispatched Order records were found in the CSV file." });
  }

  await Order.insertMany(preparedRecords);

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Dispatched Orders Management",
    action: "Import Dispatched Orders CSV",
    targetType: "Order",
    targetId: "",
    targetName: `${preparedRecords.length} Dispatched Order records`,
    details: { createdCount: preparedRecords.length }
  });

  return res.status(201).json({
    message: `CSV import completed successfully. ${preparedRecords.length} Dispatched Order record(s) created.`,
    createdCount: preparedRecords.length
  });
});

app.get("/api/super-admin/orders/dispatched/export.xlsx", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: DISPATCHED_ORDER_ROUTE_STATUSES });

  const worksheetData = result.items.map((item) => ({
    "Order Number": item.orderNumber,
    "Order Name": item.orderName,
    "Order Date & Time": item.orderDateTime,
    "Order Details Overview": item.orderDetailsOverview,
    "Customer Name": item.customerName,
    "Mobile Number": item.mobileNumber,
    "Current Status": item.currentStatus,
    "Design File Source": item.designFileSource,
    "Courier Name": item.courierName,
    "Tracking Number": item.courierTrackingNumber,
    "Dispatch Date & Time": item.dispatchDateTime,
    "Delivery Status": item.deliveryStatus,
    "Delivery Date & Time": item.deliveryDateTime
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dispatched Orders");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="dispatched-orders.xlsx"');
  return res.send(buffer);
});

app.get("/api/super-admin/orders/dispatched/export.csv", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses: DISPATCHED_ORDER_ROUTE_STATUSES });

  const header = [
    "Order Number",
    "Order Name",
    "Order Date & Time",
    "Order Details Overview",
    "Customer Name",
    "Mobile Number",
    "Current Status",
    "Design File Source",
    "Courier Name",
    "Tracking Number",
    "Dispatch Date & Time",
    "Delivery Status",
    "Delivery Date & Time"
  ];

  const csvRows = [
    header.join(","),
    ...result.items.map((item) =>
      [
        item.orderNumber,
        `"${String(item.orderName || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
        `"${String(item.customerName || "").replace(/"/g, '""')}"`,
        item.mobileNumber,
        item.currentStatus,
        `"${String(item.designFileSource || "").replace(/"/g, '""')}"`,
        `"${String(item.courierName || "").replace(/"/g, '""')}"`,
        `"${String(item.courierTrackingNumber || "").replace(/"/g, '""')}"`,
        `"${String(item.dispatchDateTime || "").replace(/"/g, '""')}"`,
        `"${String(item.deliveryStatus || "").replace(/"/g, '""')}"`,
        `"${String(item.deliveryDateTime || "").replace(/"/g, '""')}"`
      ].join(",")
    )
  ];

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="dispatched-orders.csv"');
  return res.send(csvRows.join("\n"));
});

app.get("/api/super-admin/orders/dispatched/activity/logs", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  const logs = await AuditLog.find({ module: "Dispatched Orders Management" }).sort({ createdAt: -1 }).limit(100);
  return res.json({
    items: logs.map((log) => ({
      id: log._id.toString(),
      action: log.action,
      actorName: log.actorName,
      targetName: log.targetName,
      createdAt: log.createdAt,
      details: log.details
    }))
  });
});

app.get("/api/super-admin/orders/dispatched/:orderId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Dispatched Order record ID." });
  }

  await ensureOrderNumbersForExistingOrders();

  const order = await Order.findOne({ _id: req.params.orderId, status: { $in: DISPATCHED_ORDER_ROUTE_STATUSES } });
  if (!order) {
    return res.status(404).json({ message: "Dispatched Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.patch("/api/super-admin/orders/dispatched/:orderId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Dispatched Order record ID." });
  }

  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (![ORDER_STATUS_DISPATCHED, ORDER_STATUS_COMPLETED, ORDER_STATUS_CANCELLED].includes(nextStatus)) {
    return res.status(400).json({
      message: "Dispatched Order status can only be kept as Dispatched, moved to Completed, or moved to Cancelled."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();
  const nextDeliveryStatus = String(req.body.deliveryStatus || (nextStatus === ORDER_STATUS_COMPLETED ? "delivered" : "in-transit"))
    .trim()
    .toLowerCase();

  const order = await Order.findOneAndUpdate(
    { _id: req.params.orderId, status: { $in: DISPATCHED_ORDER_ROUTE_STATUSES } },
    {
      $set: {
        status: nextStatus,
        courierName: String(req.body.courierName || "").trim() || undefined,
        courierTrackingNumber: String(req.body.courierTrackingNumber || "").trim() || undefined,
        courierTrackingUrl: String(req.body.courierTrackingUrl || "").trim() || undefined,
        dispatchNotes: String(req.body.dispatchNotes || note || "").trim() || undefined,
        deliveryStatus: nextDeliveryStatus,
        deliveryDateTime:
          nextStatus === ORDER_STATUS_COMPLETED
            ? req.body.deliveryDateTime
              ? new Date(req.body.deliveryDateTime)
              : new Date()
            : req.body.deliveryDateTime
              ? new Date(req.body.deliveryDateTime)
              : null,
        updatedAt: new Date()
      },
      $push: {
        statusHistory: {
          status: nextStatus,
          note,
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        },
        dispatchHistory: {
          status: nextStatus,
          note: String(req.body.dispatchNotes || note || "Dispatch information updated.").trim(),
          courierName: String(req.body.courierName || "").trim(),
          courierTrackingNumber: String(req.body.courierTrackingNumber || "").trim(),
          courierTrackingUrl: String(req.body.courierTrackingUrl || "").trim(),
          deliveryStatus: nextDeliveryStatus,
          eventAt: new Date(),
          updatedByRole: req.auth.role,
          updatedById: req.auth.sub,
          updatedByName: actorName
        }
      }
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Dispatched Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Dispatched Orders Management",
    action: `Update Dispatched Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      deliveryStatus: normalizeDeliveryStatusLabel(nextDeliveryStatus),
      courierName: order.courierName || "",
      courierTrackingNumber: order.courierTrackingNumber || "",
      note
    }
  });

  return res.json({
    message: `Dispatched Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

function registerScopedOrderRoutes({
  pathSegment,
  label,
  moduleName,
  allowedStatuses,
  defaultStatus,
  templateFileName,
  exportFileName,
  sampleRow,
  invalidIdMessage,
  notFoundMessage,
  allowedNextStatuses = []
}) {
  app.get(`/api/super-admin/orders/${pathSegment}`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    const result = await fetchOrderList({ ...req.query, allowedStatuses });
    return res.json(result);
  });

  app.get(`/api/super-admin/orders/${pathSegment}/template.csv`, authenticate, requireRole(ROLE_SUPER_ADMIN), (_req, res) => {
    const header = [
      "orderName",
      "customerName",
      "customerMobile",
      "orderDetailsOverview",
      "status",
      "designSubmissionSource",
      "designFileName",
      "designFileUrl",
      "referenceNo",
      "orderDateTime",
      "assignedAssociateMemberMobile",
      "assignedAdminMobile"
    ].join(",");

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${templateFileName}"`);
    return res.send(`${header}\n${sampleRow.join(",")}`);
  });

  app.post(`/api/super-admin/orders/${pathSegment}/import-csv`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    const { csvText } = req.body;

    if (!csvText || typeof csvText !== "string") {
      return res.status(400).json({ message: "CSV content is required." });
    }

    const rows = parseCsvText(csvText);
    if (rows.length < 2) {
      return res.status(400).json({ message: "CSV file must include a header row and at least one data row." });
    }

    const headers = rows[0];
    const requiredHeaders = ["orderName", "customerName", "customerMobile", "orderDetailsOverview"];
    const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
    if (missingHeaders.length > 0) {
      return res.status(400).json({ message: `Missing CSV columns: ${missingHeaders.join(", ")}` });
    }

    const validationErrors = [];
    const preparedRecords = [];
    let nextOrderNumber = await getNextOrderNumber();
    const actorName = await getActorName(req.auth);

    for (let index = 1; index < rows.length; index += 1) {
      const rowValues = rows[index];
      const rowNumber = index + 1;
      const record = Object.fromEntries(headers.map((header, headerIndex) => [header, rowValues[headerIndex] || ""]));

      const rowErrors = validateCsvOrderRecord(record, rowNumber, allowedStatuses);
      let assignedAssociateMemberId = null;
      let assignedAdminId = null;
      let placedByUserId = null;

      const normalizedMobile = String(record.customerMobile || "").replace(/\D/g, "");
      if (normalizedMobile) {
        const placedByUser = await User.findOne({ mobile: normalizedMobile }).select("_id");
        if (placedByUser) {
          placedByUserId = placedByUser._id;
        }
      }

      if (record.assignedAssociateMemberMobile) {
        const normalizedAssociateMobile = String(record.assignedAssociateMemberMobile).replace(/\D/g, "");
        const associateMember = await User.findOne({
          role: ROLE_ASSOCIATE_MEMBER,
          mobile: normalizedAssociateMobile
        }).select("_id assignedAdminId");

        if (!associateMember) {
          rowErrors.push(`Row ${rowNumber}: Assigned Associate Member mobile ${record.assignedAssociateMemberMobile} was not found.`);
        } else {
          assignedAssociateMemberId = associateMember._id;
          if (associateMember.assignedAdminId) {
            assignedAdminId = associateMember.assignedAdminId;
          }
        }
      }

      if (record.assignedAdminMobile) {
        const normalizedAdminMobile = String(record.assignedAdminMobile).replace(/\D/g, "");
        const admin = await User.findOne({ role: ROLE_ADMIN, mobile: normalizedAdminMobile }).select("_id");
        if (!admin) {
          rowErrors.push(`Row ${rowNumber}: Assigned Admin mobile ${record.assignedAdminMobile} was not found.`);
        } else {
          assignedAdminId = admin._id;
        }
      }

      if (rowErrors.length > 0) {
        validationErrors.push(...rowErrors);
        continue;
      }

      const normalizedDesignSource =
        normalizeDesignSubmissionSourceInput(record.designSubmissionSource) || DESIGN_SOURCE_ONLINE;
      const normalizedStatus = normalizeOrderStatusInput(record.status) || defaultStatus;
      const eventDateTime = record.orderDateTime ? new Date(record.orderDateTime) : new Date();
      const importNote = `${label} imported from CSV.`;

      preparedRecords.push({
        orderNumber: nextOrderNumber,
        orderName: String(record.orderName).trim(),
        customerName: String(record.customerName).trim(),
        customerMobile: normalizedMobile,
        orderDetailsOverview: String(record.orderDetailsOverview).trim(),
        orderedAt: eventDateTime,
        status: normalizedStatus,
        designSubmissionSource: normalizedDesignSource,
        designFileName: String(record.designFileName || "").trim(),
        designFileUrl: String(record.designFileUrl || "").trim(),
        referenceNo: String(record.referenceNo || "").trim(),
        placedByUserId,
        assignedAssociateMemberId,
        assignedAdminId,
        statusHistory: [
          {
            status: normalizedStatus,
            note: importNote,
            changedAt: eventDateTime,
            changedByRole: req.auth.role,
            changedById: req.auth.sub,
            changedByName: actorName
          }
        ]
      });

      nextOrderNumber += 5;
    }

    if (validationErrors.length > 0) {
      return res.status(400).json({
        message: "CSV validation failed.",
        errors: validationErrors
      });
    }

    if (preparedRecords.length === 0) {
      return res.status(400).json({ message: `No valid ${label} records were found in the CSV file.` });
    }

    await Order.insertMany(preparedRecords);

    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: moduleName,
      action: `Import ${label} CSV`,
      targetType: "Order",
      targetId: "",
      targetName: `${preparedRecords.length} ${label} records`,
      details: { createdCount: preparedRecords.length }
    });

    return res.status(201).json({
      message: `CSV import completed successfully. ${preparedRecords.length} ${label} record(s) created.`,
      createdCount: preparedRecords.length
    });
  });

  app.get(`/api/super-admin/orders/${pathSegment}/export.xlsx`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses });

    const worksheetData = result.items.map((item) => ({
      "Order Number": item.orderNumber,
      "Order Name": item.orderName,
      "Order Date & Time": item.orderDateTime,
      "Order Details Overview": item.orderDetailsOverview,
      "Customer Name": item.customerName,
      "Mobile Number": item.mobileNumber,
      "Current Status": item.currentStatus,
      "Design File Source": item.designFileSource
    }));

    const worksheet = XLSX.utils.json_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, label);
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${exportFileName}.xlsx"`);
    return res.send(buffer);
  });

  app.get(`/api/super-admin/orders/${pathSegment}/export.csv`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    const result = await fetchOrderList({ ...req.query, page: 1, limit: 10000, allowedStatuses });
    const header = [
      "Order Number",
      "Order Name",
      "Order Date & Time",
      "Order Details Overview",
      "Customer Name",
      "Mobile Number",
      "Current Status",
      "Design File Source"
    ];

    const csvRows = [
      header.join(","),
      ...result.items.map((item) =>
        [
          item.orderNumber,
          `"${String(item.orderName || "").replace(/"/g, '""')}"`,
          `"${String(item.orderDateTime || "").replace(/"/g, '""')}"`,
          `"${String(item.orderDetailsOverview || "").replace(/"/g, '""')}"`,
          `"${String(item.customerName || "").replace(/"/g, '""')}"`,
          item.mobileNumber,
          item.currentStatus,
          `"${String(item.designFileSource || "").replace(/"/g, '""')}"`
        ].join(",")
      )
    ];

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${exportFileName}.csv"`);
    return res.send(csvRows.join("\n"));
  });

  app.get(`/api/super-admin/orders/${pathSegment}/activity/logs`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
    const logs = await AuditLog.find({ module: moduleName }).sort({ createdAt: -1 }).limit(100);
    return res.json({
      items: logs.map((log) => ({
        id: log._id.toString(),
        action: log.action,
        actorName: log.actorName,
        targetName: log.targetName,
        createdAt: log.createdAt,
        details: log.details
      }))
    });
  });

  app.get(`/api/super-admin/orders/${pathSegment}/:orderId`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
      return res.status(400).json({ message: invalidIdMessage });
    }

    await ensureOrderNumbersForExistingOrders();

    const order = await Order.findOne({ _id: req.params.orderId, status: { $in: allowedStatuses } });
    if (!order) {
      return res.status(404).json({ message: notFoundMessage });
    }

    return res.json({ order: await buildOrderDetails(order) });
  });

  app.patch(`/api/super-admin/orders/${pathSegment}/:orderId/status`, authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
    if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
      return res.status(400).json({ message: invalidIdMessage });
    }

    const nextStatus = normalizeOrderStatusInput(req.body.status);
    if (!nextStatus || !allowedNextStatuses.includes(nextStatus)) {
      return res.status(400).json({
        message:
          allowedNextStatuses.length > 0
            ? `${label} status can only be updated to ${allowedNextStatuses.map((status) => normalizeOrderStatusLabel(status)).join(", ")} from this route.`
            : `${label} status updates are not available from this route.`
      });
    }

    const actorName = await getActorName(req.auth);
    const note = String(req.body.note || "").trim();

    const order = await Order.findOneAndUpdate(
      { _id: req.params.orderId, status: { $in: allowedStatuses } },
      {
        $set: {
          status: nextStatus,
          updatedAt: new Date()
        },
        $push: {
          statusHistory: {
            status: nextStatus,
            note,
            changedAt: new Date(),
            changedByRole: req.auth.role,
            changedById: req.auth.sub,
            changedByName: actorName
          }
        }
      },
      { new: true }
    );

    if (!order) {
      return res.status(404).json({ message: notFoundMessage });
    }

    await createAuditLog({
      actorRole: req.auth.role,
      actorId: req.auth.sub,
      actorName,
      module: moduleName,
      action: `Update ${label} Status to ${normalizeOrderStatusLabel(nextStatus)}`,
      targetType: "Order",
      targetId: order._id.toString(),
      targetName: order.orderName || `Order ${order.orderNumber}`,
      details: {
        orderNumber: order.orderNumber,
        nextStatus: normalizeOrderStatusLabel(nextStatus),
        note
      }
    });

    return res.json({
      message: `${label} status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
      order: await buildOrderDetails(order)
    });
  });
}

registerScopedOrderRoutes({
  pathSegment: "completed",
  label: "Completed Orders",
  moduleName: "Completed Orders Management",
  allowedStatuses: COMPLETED_ORDER_ROUTE_STATUSES,
  defaultStatus: ORDER_STATUS_COMPLETED,
  templateFileName: "completed-orders-import-template.csv",
  exportFileName: "completed-orders",
  sampleRow: [
    "Completed Retail Brochures",
    "Aarav Prints",
    "9876500051",
    "Brochure job delivered successfully to the customer",
    "Completed",
    "Online Upload",
    "completed-design.pdf",
    "https://example.com/designs/completed-design.pdf",
    "ORD-CMP-1001",
    "2026-06-20T11:30:00Z",
    "9876500002",
    "9876500003"
  ],
  invalidIdMessage: "Invalid Completed Order record ID.",
  notFoundMessage: "Completed Order record not found.",
  allowedNextStatuses: [ORDER_STATUS_COMPLETED]
});

registerScopedOrderRoutes({
  pathSegment: "cancelled",
  label: "Cancelled Orders",
  moduleName: "Cancelled Orders Management",
  allowedStatuses: CANCELLED_ORDER_ROUTE_STATUSES,
  defaultStatus: ORDER_STATUS_CANCELLED,
  templateFileName: "cancelled-orders-import-template.csv",
  exportFileName: "cancelled-orders",
  sampleRow: [
    "Cancelled Box Packaging",
    "Nexa Packaging",
    "9876500052",
    "Order cancelled before production due to customer request",
    "Cancelled",
    "Send Design via Email",
    "",
    "",
    "ORD-CAN-1001",
    "2026-06-20T12:00:00Z",
    "9876500002",
    "9876500003"
  ],
  invalidIdMessage: "Invalid Cancelled Order record ID.",
  notFoundMessage: "Cancelled Order record not found.",
  allowedNextStatuses: [ORDER_STATUS_CANCELLED]
});

registerScopedOrderRoutes({
  pathSegment: "improper",
  label: "Improper Orders",
  moduleName: "Improper Orders Management",
  allowedStatuses: IMPROPER_ORDER_ROUTE_STATUSES,
  defaultStatus: ORDER_STATUS_IMPROPER,
  templateFileName: "improper-orders-import-template.csv",
  exportFileName: "improper-orders",
  sampleRow: [
    "Improper Visiting Cards",
    "City Offset Works",
    "9876500053",
    "Incorrect specifications submitted; requires manual correction",
    "Improper",
    "Online Upload",
    "improper-design.pdf",
    "https://example.com/designs/improper-design.pdf",
    "ORD-IMP-1001",
    "2026-06-20T13:00:00Z",
    "9876500002",
    "9876500003"
  ],
  invalidIdMessage: "Invalid Improper Order record ID.",
  notFoundMessage: "Improper Order record not found.",
  allowedNextStatuses: [ORDER_STATUS_IMPROPER]
});

registerScopedOrderRoutes({
  pathSegment: "rejected",
  label: "Rejected Orders",
  moduleName: "Rejected Orders Management",
  allowedStatuses: REJECTED_ORDER_ROUTE_STATUSES,
  defaultStatus: ORDER_STATUS_REJECTED,
  templateFileName: "rejected-orders-import-template.csv",
  exportFileName: "rejected-orders",
  sampleRow: [
    "Rejected Promo Flyers",
    "Classic Media House",
    "9876500054",
    "Rejected after final review because mandatory assets were missing",
    "Rejected",
    "Send Design via Email",
    "",
    "",
    "ORD-REJ-1001",
    "2026-06-20T14:00:00Z",
    "9876500002",
    "9876500003"
  ],
  invalidIdMessage: "Invalid Rejected Order record ID.",
  notFoundMessage: "Rejected Order record not found.",
  allowedNextStatuses: [ORDER_STATUS_REJECTED]
});

app.get("/api/super-admin/orders/:orderId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Order record ID." });
  }

  await ensureOrderNumbersForExistingOrders();

  const order = await Order.findById(req.params.orderId);
  if (!order) {
    return res.status(404).json({ message: "Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.patch("/api/super-admin/orders/:orderId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Order record ID." });
  }

  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (!nextStatus) {
    return res.status(400).json({
      message: "Status must be Pending, Printing, Packaging, Dispatched, Cancelled, Improper, Completed, or Rejected."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();
  const setPayload = {
    status: nextStatus,
    updatedAt: new Date()
  };
  const pushPayload = {
    statusHistory: {
      status: nextStatus,
      note,
      changedAt: new Date(),
      changedByRole: req.auth.role,
      changedById: req.auth.sub,
      changedByName: actorName
    }
  };

  if (nextStatus === ORDER_STATUS_DISPATCHED) {
    setPayload.dispatchDateTime = req.body.dispatchDateTime ? new Date(req.body.dispatchDateTime) : new Date();
    setPayload.courierName = String(req.body.courierName || "").trim();
    setPayload.courierTrackingNumber = String(req.body.courierTrackingNumber || "").trim();
    setPayload.courierTrackingUrl = String(req.body.courierTrackingUrl || "").trim();
    setPayload.dispatchNotes = String(req.body.dispatchNotes || note || "Order dispatched.").trim();
    setPayload.deliveryStatus = String(req.body.deliveryStatus || "dispatched").trim().toLowerCase();

    pushPayload.dispatchHistory = {
      status: nextStatus,
      note: String(req.body.dispatchNotes || note || "Order dispatched.").trim(),
      courierName: String(req.body.courierName || "").trim(),
      courierTrackingNumber: String(req.body.courierTrackingNumber || "").trim(),
      courierTrackingUrl: String(req.body.courierTrackingUrl || "").trim(),
      deliveryStatus: String(req.body.deliveryStatus || "dispatched").trim().toLowerCase(),
      eventAt: req.body.dispatchDateTime ? new Date(req.body.dispatchDateTime) : new Date(),
      updatedByRole: req.auth.role,
      updatedById: req.auth.sub,
      updatedByName: actorName
    };
  }

  if (nextStatus === ORDER_STATUS_COMPLETED) {
    setPayload.deliveryStatus = String(req.body.deliveryStatus || "delivered").trim().toLowerCase();
    setPayload.deliveryDateTime = req.body.deliveryDateTime ? new Date(req.body.deliveryDateTime) : new Date();
  }

  const order = await Order.findByIdAndUpdate(
    req.params.orderId,
    {
      $set: setPayload,
      $push: pushPayload
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "All Orders Management",
    action: `Update Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      note,
      courierName: order.courierName || "",
      courierTrackingNumber: order.courierTrackingNumber || "",
      deliveryStatus: normalizeDeliveryStatusLabel(order.deliveryStatus)
    }
  });

  return res.json({
    message: `Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

function getAdminModuleOrderStatuses(view) {
  switch (String(view || "").trim().toLowerCase()) {
    case "new":
      return [ORDER_STATUS_PENDING];
    case "pending":
      return [ORDER_STATUS_PENDING];
    case "printing":
      return [ORDER_STATUS_PRINTING];
    case "packaging":
      return [ORDER_STATUS_PACKAGING];
    case "dispatch":
      return [ORDER_STATUS_DISPATCHED];
    case "completed":
      return [ORDER_STATUS_COMPLETED];
    case "improper":
      return [ORDER_STATUS_IMPROPER];
    case "cancelled":
      return [ORDER_STATUS_CANCELLED];
    case "rejected":
      return [ORDER_STATUS_REJECTED];
    default:
      return [];
  }
}

function getAssociateMemberOrderStatuses(view) {
  switch (String(view || "").trim().toLowerCase()) {
    case "pending":
      return [ORDER_STATUS_PENDING];
    case "printing":
      return [ORDER_STATUS_PRINTING, ORDER_STATUS_PACKAGING, ORDER_STATUS_DISPATCHED];
    case "completed":
      return [ORDER_STATUS_COMPLETED];
    default:
      return [];
  }
}

async function getAssociateMemberRecord(memberId) {
  return User.findOne({ _id: memberId, role: ROLE_ASSOCIATE_MEMBER }).select(
    "ownerName mobile email businessName country state district city pinCode address services status associateMemberAccessEnabled assignedAdminId gstNumber referenceCode createdAt updatedAt lastLoginAt lastActivityAt"
  );
}

function mapAssociateServiceCard(service) {
  const normalized = String(service || "").trim();
  const labelMap = {
    printingServices: "Printing Services",
    expo: "Expo Promotions",
    magazine: "Magazine Listing",
    advertiser: "Advertising Support"
  };

  return {
    title: labelMap[normalized] || normalized || "Printing Service",
    description: "This service is active for your associate member account and can be used in order-booking workflows.",
    meta: "Service enabled"
  };
}

async function buildAssociateMemberModuleBootstrap(memberId) {
  await ensureOrderNumbersForExistingOrders();

  const associateMember = await getAssociateMemberRecord(memberId);
  if (!associateMember) {
    return null;
  }

  const [totalOrders, pendingOrders, activeProductionOrders, completedOrders, topUps, recentOrders, walletBalance, assignedAdmin] = await Promise.all([
    Order.countDocuments({ placedByUserId: memberId }),
    Order.countDocuments({ placedByUserId: memberId, status: ORDER_STATUS_PENDING }),
    Order.countDocuments({ placedByUserId: memberId, status: { $in: [ORDER_STATUS_PRINTING, ORDER_STATUS_PACKAGING, ORDER_STATUS_DISPATCHED] } }),
    Order.countDocuments({ placedByUserId: memberId, status: ORDER_STATUS_COMPLETED }),
    TopUpRequest.find({ requestedByUserId: memberId }).sort({ createdAt: -1 }).limit(50),
    Order.find({ placedByUserId: memberId }).sort({ updatedAt: -1 }).limit(6).select("orderNumber orderName status orderedAt"),
    getAssociateWalletBalance(memberId),
    associateMember.assignedAdminId ? User.findById(associateMember.assignedAdminId).select("ownerName businessName mobile email") : null
  ]);

  const pendingTopUps = topUps.filter((request) => request.status === "pending").length;

  return {
    header: {
      systemName: "Printing Services Division",
      systemDescription: "Associate member portal based on the existing dashboard layout, responsive tables, cards, and modal workflow.",
      memberId: `ASM-${String(associateMember.mobile || "0000").slice(-4).padStart(4, "0")}`,
      walletBalance: walletBalance.toFixed(2)
    },
    footer: {
      companyName: "Printers Club of India Limited",
      companyDescription:
        "Dedicated to the continuous development and modernization of the printing industry in India. Providing quality services and a unified platform for printers nationwide.",
      contactEmail: "support@printersclub.in",
      contactPhone: "+91 99758 13249",
      contactAddress: "Plot No. 57, Jhotwara Industrial Area, Jaipur, Rajasthan",
      quickLinks: [
        { label: "Dashboard", href: "/dashboard/associate-member" },
        { label: "Book Order", href: "/dashboard/associate-member/book-order" },
        { label: "My Orders", href: "/dashboard/associate-member/orders/all" },
        { label: "Wallet", href: "/dashboard/associate-member/wallet" }
      ],
      termsLink: "Terms_And_Conditions.aspx?type=new"
    },
    profileDefaults: {
      ownerName: associateMember.ownerName || "",
      businessName: associateMember.businessName || "",
      email: associateMember.email || "",
      mobile: associateMember.mobile || "",
      country: associateMember.country || "",
      state: associateMember.state || "",
      district: associateMember.district || "",
      city: associateMember.city || "",
      pinCode: associateMember.pinCode || "",
      address: associateMember.address || ""
    },
    profile: {
      ownerName: associateMember.ownerName || "",
      businessName: associateMember.businessName || "",
      email: associateMember.email || "",
      mobile: associateMember.mobile || "",
      country: associateMember.country || "",
      state: associateMember.state || "",
      district: associateMember.district || "",
      city: associateMember.city || "",
      pinCode: associateMember.pinCode || "",
      address: associateMember.address || "",
      gstNumber: associateMember.gstNumber || "",
      referenceCode: associateMember.referenceCode || "",
      accountStatus: normalizeAdminStatusLabel(associateMember.status),
      associateMemberAccess: associateMember.associateMemberAccessEnabled ? "Enabled" : "Disabled",
      registrationDate: associateMember.createdAt || null,
      lastUpdatedAt: associateMember.updatedAt || null,
      lastLoginAt: associateMember.lastLoginAt || null,
      lastActivityAt: associateMember.lastActivityAt || null,
      walletBalance: walletBalance.toFixed(2),
      assignedAdminName: assignedAdmin?.ownerName || "",
      assignedAdminBusinessName: assignedAdmin?.businessName || "",
      assignedAdminMobile: assignedAdmin?.mobile || "",
      assignedAdminEmail: assignedAdmin?.email || ""
    },
    dashboardCards: [
      { title: "Total Orders", value: totalOrders, note: "All orders raised from your account.", iconKey: "my-orders" },
      { title: "Pending Orders", value: pendingOrders, note: "Orders currently waiting for production.", iconKey: "my-orders" },
      { title: "In Production", value: activeProductionOrders, note: "Printing, packaging, and dispatch pipeline.", iconKey: "my-orders" },
      { title: "Wallet Requests", value: pendingTopUps, note: "Top-up requests awaiting review.", iconKey: "wallet" }
    ],
    serviceCards: (associateMember.services || []).length
      ? associateMember.services.map(mapAssociateServiceCard)
      : [
          {
            title: "Printing Services",
            description: "Submit new print jobs and manage order workflows from your dashboard.",
            meta: "Default service"
          }
        ],
    recentOrderColumns: [
      { key: "orderNumber", label: "Order No." },
      { key: "orderName", label: "Order Name" },
      { key: "status", label: "Status" },
      { key: "orderedAt", label: "Booked On" }
    ],
    recentOrders: recentOrders.map((order) => ({
      id: order._id.toString(),
      orderNumber: order.orderNumber || "--",
      orderName: order.orderName || "--",
      status: normalizeOrderStatusLabel(order.status),
      orderedAt: formatDateTime(order.orderedAt)
    }))
  };
}

async function buildAssociateMemberSectionData(section, view, memberId) {
  if (section === "my-orders") {
    await ensureOrderNumbersForExistingOrders();

    const query = { placedByUserId: memberId };
    const statuses = getAssociateMemberOrderStatuses(view);
    if (statuses.length > 0) {
      query.status = { $in: statuses };
    }

    const orders = await Order.find(query)
      .sort({ createdAt: -1, orderedAt: -1 })
      .limit(100)
      .select(
        "orderNumber orderName customerName customerMobile status orderDetailsOverview bagName printSide quantity bagSize bagColor textColorType textColors printingPress privacy deliveryOption fileOption sellingPrice remark pressline referenceNo designSubmissionSource designFileName designFileType designFileUrl orderedAt basePayableAmount pdfDiscountAmount walletDebitAmount createdAt"
      );

    return {
      meta: {
        columns: [
          { key: "orderNumber", label: "Order No." },
          { key: "dateTime", label: "Date & Time" },
          { key: "orderName", label: "Order Name" },
          { key: "orderDetail", label: "Order Summary" },
          { key: "status", label: "Status" },
          { key: "fileType", label: "File Type" }
        ]
      },
      summary: {
        total: orders.length,
        pending: orders.filter((order) => order.status === ORDER_STATUS_PENDING).length,
        completed: orders.filter((order) => order.status === ORDER_STATUS_COMPLETED).length
      },
      items: orders.map((order) => ({
        id: order._id.toString(),
        orderNumber: order.orderNumber || "--",
        dateTime: formatDateTime(order.orderedAt || order.createdAt),
        orderName: order.orderName || "--",
        orderDetail: buildRecentOrderSummary(order),
        status: normalizeOrderStatusLabel(order.status),
        fileType: getRecentOrderFileType(order),
        bagName: order.bagName || "--",
        printSide: order.printSide || "--",
        quantity: Number(order.quantity || 0),
        bagSize: order.bagSize || "--",
        bagColor: order.bagColor || "--",
        textColorType: order.textColorType || "--",
        textColors: Array.isArray(order.textColors) ? order.textColors : [],
        customerName: order.customerName || "--",
        customerMobile: order.customerMobile || "--",
        orderDetailsOverview: order.orderDetailsOverview || "--",
        printingPress: order.printingPress || "--",
        privacy: order.privacy || "--",
        deliveryOption: order.deliveryOption || "--",
        fileOption: order.fileOption || "--",
        sellingPrice: Number(order.sellingPrice || 0),
        remark: order.remark || "",
        pressline: order.pressline || "--",
        referenceNo: order.referenceNo || "--",
        designSubmissionSource: order.designSubmissionSource || "",
        designFileName: order.designFileName || "",
        designFileType: order.designFileType || "",
        designFileUrl: order.designFileUrl || "",
        basePayableAmount: Number(order.basePayableAmount || 0),
        pdfDiscountAmount: Number(order.pdfDiscountAmount || 0),
        walletDebitAmount: Number(order.walletDebitAmount || 0)
      }))
    };
  }

  if (section === "wallet") {
    return buildWalletHistoryForActor({ userId: memberId, role: ROLE_ASSOCIATE_MEMBER });
  }

  if (section === "notifications") {
    const [orders, topUps] = await Promise.all([
      Order.find({ placedByUserId: memberId }).sort({ updatedAt: -1 }).limit(15).select("orderNumber orderName status updatedAt"),
      TopUpRequest.find({ requestedByUserId: memberId }).sort({ updatedAt: -1 }).limit(10).select("status amount updatedAt")
    ]);

    const items = [
      ...orders.map((order) => ({
        id: `order-${order._id.toString()}`,
        title: `Order ${order.orderNumber || "--"}`,
        message: `${order.orderName || "Order"} is now ${normalizeOrderStatusLabel(order.status)}.`,
        type: "Order Update",
        status: "Unread",
        date: formatDateTime(order.updatedAt),
        sortValue: order.updatedAt
      })),
      ...topUps.map((request) => ({
        id: `topup-${request._id.toString()}`,
        title: "Wallet Request",
        message: `Top-up request of Rs. ${Number(request.amount || 0).toFixed(2)} is ${String(request.status || "").toLowerCase()}.`,
        type: "Wallet",
        status: "Unread",
        date: formatDateTime(request.updatedAt),
        sortValue: request.updatedAt
      }))
    ]
      .sort((left, right) => new Date(right.sortValue).getTime() - new Date(left.sortValue).getTime())
      .map(({ sortValue, ...item }) => item);

    return {
      meta: {
        columns: [
          { key: "title", label: "Notification" },
          { key: "message", label: "Message" },
          { key: "type", label: "Type" },
          { key: "status", label: "Status" },
          { key: "date", label: "Updated At" }
        ]
      },
      summary: {
        total: items.length
      },
      items
    };
  }

  if (section === "profile") {
    const associateMember = await getAssociateMemberRecord(memberId);

    return {
      meta: {
        columns: [
          { key: "field", label: "Field" },
          { key: "value", label: "Value" }
        ]
      },
      summary: {
        total: associateMember ? 9 : 0
      },
      items: associateMember
        ? [
            { id: "name", field: "Your Name", value: associateMember.ownerName },
            { id: "business", field: "Business / Firm", value: associateMember.businessName },
            { id: "mobile", field: "WhatsApp Number", value: associateMember.mobile },
            { id: "email", field: "Email Address", value: associateMember.email },
            { id: "country", field: "Country", value: associateMember.country },
            { id: "state", field: "State", value: associateMember.state },
            { id: "district", field: "District", value: associateMember.district },
            { id: "address", field: "Full Address", value: buildAddress(associateMember) },
            { id: "access", field: "Associate Access", value: associateMember.associateMemberAccessEnabled ? "Enabled" : "Disabled" }
          ]
        : []
    };
  }

  return {
    meta: { columns: [] },
    summary: { total: 0 },
    items: []
  };
}

app.get("/api/associate-member-module/bootstrap", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  const data = await buildAssociateMemberModuleBootstrap(req.auth.sub);

  if (!data) {
    return res.status(404).json({ message: "Associate Member record not found." });
  }

  return res.json(data);
});

app.get("/api/associate-member-module/section/:section", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  return res.json(await buildAssociateMemberSectionData(req.params.section, req.query.view, req.auth.sub));
});

app.get("/api/associate-member-module/orders/search", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  const searchType = String(req.query.searchType || "").trim().toLowerCase();
  const orderNumber = String(req.query.orderNumber || "").trim();
  const status = String(req.query.status || "").trim();
  const fromDate = String(req.query.fromDate || "").trim();
  const toDate = String(req.query.toDate || "").trim();

  const queryPayload = {
    placedByUserId: req.auth.sub,
    limit: 100,
    sortBy: "orderedAt",
    sortOrder: "desc"
  };

  if (searchType === "order-number") {
    queryPayload.orderNumber = orderNumber;
  }

  if (searchType === "order-stage") {
    queryPayload.status = status;
  }

  if (searchType === "order-date") {
    queryPayload.fromDate = fromDate;
    queryPayload.toDate = toDate || fromDate;
  }

  const result = await fetchOrderList(queryPayload);
  const items = Array.isArray(result?.items) ? result.items.map(mapAssociateSearchOrderRow) : [];

  return res.json({
    meta: {
      columns: [
        { key: "orderNumber", label: "Order No." },
        { key: "dateTime", label: "Date & Time" },
        { key: "orderName", label: "Order Name" },
        { key: "orderDetail", label: "Order Summary" },
        { key: "status", label: "Current Status" },
        { key: "fileType", label: "File Type" }
      ]
    },
    summary: {
      total: items.length,
      pending: items.filter((order) => String(order.currentStatusValue || "").toLowerCase() === ORDER_STATUS_PENDING).length,
      completed: items.filter((order) => String(order.currentStatusValue || "").toLowerCase() === ORDER_STATUS_COMPLETED).length
    },
    items
  });
});

app.get("/api/associate-member-module/orders/:orderId", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Order record ID." });
  }

  const order = await Order.findOne({ _id: req.params.orderId, placedByUserId: req.auth.sub });
  if (!order) {
    return res.status(404).json({ message: "Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

async function createManualOrderForRole(req, res, role) {
  const orderName = String(req.body.orderName || "").trim();
  const customerName = String(req.body.customerName || "").trim();
  const orderDetailsOverview = String(req.body.orderDetailsOverview || "").trim();
  const designSubmissionSource = String(req.body.designSubmissionSource || DESIGN_SOURCE_ONLINE).trim().toLowerCase();
  const designFileName = String(req.body.designFileName || "").trim();
  const designFileType = String(req.body.designFileType || "").trim();
  const designFileUrl = String(req.body.designFileUrl || "").trim();
  const bagName = String(req.body.bagName || "").trim();
  const printSide = String(req.body.printSide || "").trim();
  const quantity = Number(req.body.quantity || 0);
  const bagSize = String(req.body.bagSize || "").trim();
  const bagColor = String(req.body.bagColor || "").trim();
  const textColorType = String(req.body.textColorType || "").trim();
  const textColors = Array.isArray(req.body.textColors) ? req.body.textColors.map((color) => String(color || "").trim()).filter(Boolean) : [];
  const printingPress = String(req.body.printingPress || "").trim();
  const privacy = String(req.body.privacy || "").trim();
  const deliveryOption = String(req.body.deliveryOption || "").trim();
  const fileOption = String(req.body.fileOption || "").trim();
  const sellingPrice = Number(req.body.sellingPrice || 0);
  const remark = String(req.body.remark || "").trim();
  const pressline = String(req.body.pressline || "").trim();
  const referenceNo = String(req.body.referenceNo || "").trim();
  const isUrgent = Boolean(req.body.isUrgent);
  const basePayableAmount = Number(req.body.basePayableAmount || req.body.payableAmount || 0);
  const normalizedDesignFileType = String(designFileType || (designFileName.includes(".") ? designFileName.split(".").pop() : ""))
    .trim()
    .toLowerCase();
  const pdfDiscountAmount = designSubmissionSource === DESIGN_SOURCE_ONLINE && normalizedDesignFileType === "pdf" ? 10 : 0;
  const payableAmount = Math.max(0, Number((basePayableAmount - pdfDiscountAmount).toFixed(2)));
  const actorName = await getActorName(req.auth);

  let actorUser = null;
  let normalizedCountry = String(req.auth.country || "India").trim();
  let placedByUserId = null;
  let assignedAssociateMemberId = null;
  let assignedAdminId = null;
  let moduleName = "Order Management";
  let createdNote = "Order created successfully.";
  let walletBalance = null;

  if (role === ROLE_ASSOCIATE_MEMBER) {
    actorUser = await getAssociateMemberRecord(req.auth.sub);
    if (!actorUser) {
      return res.status(404).json({ message: "Associate Member record not found." });
    }

    normalizedCountry = actorUser.country || normalizedCountry;
    placedByUserId = actorUser._id;
    assignedAssociateMemberId = actorUser._id;
    assignedAdminId = actorUser.assignedAdminId || null;
    moduleName = "Associate Member Module";
    createdNote = "Order created by associate member.";
    walletBalance = await getAssociateWalletBalance(actorUser._id);
  } else if (role === ROLE_ADMIN) {
    actorUser = await User.findOne({ _id: req.auth.sub, role: ROLE_ADMIN }).select("_id ownerName mobile email businessName country");
    if (!actorUser) {
      return res.status(404).json({ message: "Admin record not found." });
    }

    normalizedCountry = actorUser.country || normalizedCountry;
    placedByUserId = actorUser._id;
    assignedAdminId = actorUser._id;
    moduleName = "Admin Module";
    createdNote = "Order created by admin.";
  } else if (role === ROLE_SUPER_ADMIN) {
    moduleName = "Super Admin Module";
    createdNote = "Order created by super admin.";
  }

  const customerMobile = normalizeMobileNumber(normalizedCountry, req.body.customerMobile);

  if (!orderName || !customerName || !customerMobile || !orderDetailsOverview) {
    return res.status(400).json({ message: "Order name, customer name, mobile number, and order details are required." });
  }

  if (!Number.isFinite(basePayableAmount) || basePayableAmount <= 0) {
    return res.status(400).json({ message: "Enter a valid payable amount." });
  }

  if (!/^\d{10,15}$/.test(customerMobile)) {
    return res.status(400).json({ message: "Enter a valid customer mobile number." });
  }

  if (![DESIGN_SOURCE_ONLINE, DESIGN_SOURCE_EMAIL].includes(designSubmissionSource)) {
    return res.status(400).json({ message: "Invalid design submission source." });
  }

  if (!bagName || !printSide || !Number.isFinite(quantity) || quantity <= 0 || !bagSize || !bagColor || !textColorType || !textColors.length || !printingPress || !privacy || !deliveryOption || !fileOption || !Number.isFinite(sellingPrice) || sellingPrice <= 0 || !pressline) {
    return res.status(400).json({ message: "Complete order details are required." });
  }

  if (walletBalance !== null && walletBalance < payableAmount) {
    return res.status(400).json({
      message: "Insufficient balance in wallet.",
      walletBalance: walletBalance.toFixed(2),
      payableAmount: payableAmount.toFixed(2)
    });
  }

  const walletDebitAmount = Number(payableAmount.toFixed(2));
  await ensureOrderNumbersForExistingOrders();
  const maxOrder = await Order.findOne({}).sort({ orderNumber: -1 }).select("orderNumber");
  const nextOrderNumber = Math.max(Number(maxOrder?.orderNumber || 1000) + 1, 1001);

  const order = await Order.create({
    orderNumber: nextOrderNumber,
    orderName,
    customerName,
    customerMobile,
    orderDetailsOverview,
    bagName,
    printSide,
    quantity,
    bagSize,
    bagColor,
    textColorType,
    textColors,
    printingPress,
    privacy,
    deliveryOption,
    fileOption,
    sellingPrice,
    remark,
    pressline,
    status: ORDER_STATUS_PENDING,
    designSubmissionSource,
    designFileName,
    designFileType,
    designFileUrl,
    isUrgent,
    referenceNo,
    basePayableAmount: Number(basePayableAmount.toFixed(2)),
    pdfDiscountAmount,
    walletDebitAmount,
    placedByUserId,
    assignedAssociateMemberId,
    assignedAdminId,
    statusHistory: [
      {
        status: ORDER_STATUS_PENDING,
        note: createdNote,
        changedAt: new Date(),
        changedByRole: req.auth.role,
        changedById: req.auth.sub,
        changedByName: actorName
      }
    ]
  });

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: moduleName,
    action: "Book Order",
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(order.status),
      pdfDiscountAmount,
      walletDebitAmount
    }
  });

  const responsePayload = {
    message: `Order ${order.orderNumber} booked successfully.`,
    walletDebitedAmount: walletDebitAmount.toFixed(2)
  };

  if (walletBalance !== null) {
    const remainingWalletBalance = Math.max(0, Number((walletBalance - walletDebitAmount).toFixed(2)));
    responsePayload.walletBalance = remainingWalletBalance.toFixed(2);
  }

  return res.status(201).json(responsePayload);
}

app.post("/api/associate-member-module/orders", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  return createManualOrderForRole(req, res, ROLE_ASSOCIATE_MEMBER);
});

app.get("/api/admin-module/orders/recent", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const orders = await Order.find({ assignedAdminId: req.auth.sub })
    .sort({ updatedAt: -1 })
    .limit(Math.min(Math.max(Number(req.query.limit) || 5, 1), 20))
    .select(
      "orderNumber orderName customerName customerMobile status orderDetailsOverview bagName printSide quantity bagSize bagColor textColorType textColors printingPress privacy deliveryOption fileOption sellingPrice remark pressline referenceNo designSubmissionSource designFileName designFileType designFileUrl orderedAt basePayableAmount pdfDiscountAmount walletDebitAmount createdAt"
    );

  return res.json({
    items: orders.map(mapRecentOrderRow)
  });
});

app.post("/api/admin-module/orders", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return createManualOrderForRole(req, res, ROLE_ADMIN);
});

app.get("/api/super-admin/orders/recent", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  const orders = await Order.find({})
    .sort({ updatedAt: -1 })
    .limit(Math.min(Math.max(Number(req.query.limit) || 5, 1), 20))
    .select(
      "orderNumber orderName customerName customerMobile status orderDetailsOverview bagName printSide quantity bagSize bagColor textColorType textColors printingPress privacy deliveryOption fileOption sellingPrice remark pressline referenceNo designSubmissionSource designFileName designFileType designFileUrl orderedAt basePayableAmount pdfDiscountAmount walletDebitAmount createdAt"
    );

  return res.json({
    items: orders.map(mapRecentOrderRow)
  });
});

app.post("/api/super-admin/orders", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return createManualOrderForRole(req, res, ROLE_SUPER_ADMIN);
});

async function createWalletTopUpForRole(req, res, role, moduleName) {
  const amount = Number(req.body.amount || 0);
  const remarks = String(req.body.remarks || "").trim();

  if (!Number.isFinite(amount) || amount <= 0) {
    return res.status(400).json({ message: "Enter a valid top-up amount greater than zero." });
  }

  const topUpActor = buildTopUpActor(req.auth);
  const actorName = await getActorName(req.auth);
  const topUpRequest = await TopUpRequest.create({
    ...topUpActor,
    amount,
    status: "approved",
    remarks,
    reviewedByRole: req.auth.role,
    reviewedById: req.auth.sub,
    reviewedByName: actorName,
    reviewedAt: new Date(),
    decisionNote: remarks
  });

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: moduleName,
    action: "Add Money To Wallet",
    targetType: "Top-Up Request",
    targetId: topUpRequest._id.toString(),
    targetName: actorName,
    details: {
      amount,
      remarks
    }
  });

  const walletBalance = await getWalletBalanceForActor({
    userId: req.auth.role === ROLE_SUPER_ADMIN ? "" : req.auth.sub,
    role,
    externalKey: role === ROLE_SUPER_ADMIN ? "super-admin" : ""
  });
  return res.status(201).json({
    message: "Money added to wallet successfully.",
    walletBalance: walletBalance.toFixed(2)
  });
}

app.post("/api/associate-member-module/wallet/top-up", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  return createWalletTopUpForRole(req, res, ROLE_ASSOCIATE_MEMBER, "Associate Member Wallet");
});

app.get("/api/associate-member-module/wallet/history", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  return res.json(await buildWalletHistoryForActor({ userId: req.auth.sub, role: ROLE_ASSOCIATE_MEMBER }));
});

app.get("/api/admin-module/wallet/history", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return res.json(await buildWalletHistoryForActor({ userId: req.auth.sub, role: ROLE_ADMIN }));
});

app.get("/api/super-admin/wallet/history", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  return res.json(await buildWalletHistoryForActor({ role: ROLE_SUPER_ADMIN, externalKey: "super-admin" }));
});

app.get("/api/super-admin/wallet/transactions", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (_req, res) => {
  return res.json(await listWalletTransactionsForSuperAdmin());
});

app.get("/api/admin-module/wallet/requests/:requestId", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.requestId)) {
    return res.status(400).json({ message: "Invalid wallet request ID." });
  }

  const assignedMemberIds = await User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: req.auth.sub }).distinct("_id");
  const walletRequest = await TopUpRequest.findOne({
    _id: req.params.requestId,
    requestedByUserId: { $in: assignedMemberIds }
  }).populate("requestedByUserId", "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt");

  if (!walletRequest) {
    return res.status(404).json({ message: "Wallet request not found." });
  }

  return res.json({
    walletRequest: await buildWalletRequestDetails(walletRequest)
  });
});

app.patch("/api/admin-module/wallet/requests/:requestId/status", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.requestId)) {
    return res.status(400).json({ message: "Invalid wallet request ID." });
  }

  const nextStatus = String(req.body.status || "").trim().toLowerCase();
  if (!["approved", "rejected"].includes(nextStatus)) {
    return res.status(400).json({ message: "Wallet request status must be Approved or Rejected." });
  }

  const assignedMemberIds = await User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: req.auth.sub }).distinct("_id");
  const actorName = await getActorName(req.auth);
  const decisionNote = String(req.body.note || "").trim();

  const walletRequest = await TopUpRequest.findOneAndUpdate(
    {
      _id: req.params.requestId,
      requestedByUserId: { $in: assignedMemberIds }
    },
    {
      $set: {
        status: nextStatus,
        reviewedByRole: req.auth.role,
        reviewedById: req.auth.sub,
        reviewedByName: actorName,
        reviewedAt: new Date(),
        decisionNote
      }
    },
    { new: true }
  ).populate("requestedByUserId", "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt");

  if (!walletRequest) {
    return res.status(404).json({ message: "Wallet request not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Wallet Management",
    action: nextStatus === "approved" ? "Approve Wallet Request" : "Reject Wallet Request",
    targetType: "Top-Up Request",
    targetId: walletRequest._id.toString(),
    targetName: walletRequest.requestedByDisplayName || walletRequest.requestedByUserId?.ownerName || "Wallet Request",
    details: {
      amount: Number(walletRequest.amount || 0).toFixed(2),
      nextStatus: normalizeTopUpStatusLabel(nextStatus),
      decisionNote
    }
  });

  return res.json({
    message: `Wallet request ${nextStatus} successfully.`,
    walletRequest: await buildWalletRequestDetails(walletRequest)
  });
});

app.post("/api/admin-module/wallet/top-up", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return createWalletTopUpForRole(req, res, ROLE_ADMIN, "Admin Wallet");
});

app.post("/api/super-admin/wallet/top-up", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return createWalletTopUpForRole(req, res, ROLE_SUPER_ADMIN, "Super Admin Wallet");
});

app.get("/api/super-admin/wallet/requests", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return res.json(await listWalletRequestsForSuperAdmin(req.query.view || "transactions"));
});

app.get("/api/super-admin/wallet/requests/:requestId", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.requestId)) {
    return res.status(400).json({ message: "Invalid wallet request ID." });
  }

  const walletRequest = await TopUpRequest.findById(req.params.requestId).populate(
    "requestedByUserId",
    "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt"
  );

  if (!walletRequest) {
    return res.status(404).json({ message: "Wallet request not found." });
  }

  return res.json({
    walletRequest: await buildWalletRequestDetails(walletRequest)
  });
});

app.patch("/api/super-admin/wallet/requests/:requestId/status", authenticate, requireRole(ROLE_SUPER_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.requestId)) {
    return res.status(400).json({ message: "Invalid wallet request ID." });
  }

  const nextStatus = String(req.body.status || "").trim().toLowerCase();
  if (!["approved", "rejected"].includes(nextStatus)) {
    return res.status(400).json({ message: "Wallet request status must be Approved or Rejected." });
  }

  const actorName = await getActorName(req.auth);
  const decisionNote = String(req.body.note || "").trim();

  const walletRequest = await TopUpRequest.findByIdAndUpdate(
    req.params.requestId,
    {
      $set: {
        status: nextStatus,
        reviewedByRole: req.auth.role,
        reviewedById: req.auth.sub,
        reviewedByName: actorName,
        reviewedAt: new Date(),
        decisionNote
      }
    },
    { new: true }
  ).populate("requestedByUserId", "ownerName mobile email businessName address country state district city pinCode gstNumber role status createdAt");

  if (!walletRequest) {
    return res.status(404).json({ message: "Wallet request not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Super Admin Wallet Management",
    action: nextStatus === "approved" ? "Approve Wallet Request" : "Reject Wallet Request",
    targetType: "Top-Up Request",
    targetId: walletRequest._id.toString(),
    targetName: walletRequest.requestedByDisplayName || walletRequest.requestedByUserId?.ownerName || "Wallet Request",
    details: {
      amount: Number(walletRequest.amount || 0).toFixed(2),
      nextStatus: normalizeTopUpStatusLabel(nextStatus),
      decisionNote
    }
  });

  return res.json({
    message: `Wallet request ${nextStatus} successfully.`,
    walletRequest: await buildWalletRequestDetails(walletRequest)
  });
});

app.patch("/api/associate-member-module/profile", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  const associateMember = await getAssociateMemberRecord(req.auth.sub);
  if (!associateMember) {
    return res.status(404).json({ message: "Associate Member record not found." });
  }

  const ownerName = String(req.body.ownerName || "").trim();
  const businessName = String(req.body.businessName || "").trim();
  const email = String(req.body.email || "").trim().toLowerCase();
  const country = String(req.body.country || "").trim();
  const mobile = normalizeMobileNumber(country, req.body.mobile);
  const state = String(req.body.state || "").trim();
  const district = String(req.body.district || "").trim();
  const city = String(req.body.city || "").trim();
  const pinCode = String(req.body.pinCode || "").trim();
  const address = String(req.body.address || "").trim();

  if (!ownerName || !businessName || !email || !country || !mobile || !state || !district || !city || !pinCode || !address) {
    return res.status(400).json({ message: "Please complete all required profile fields." });
  }

  const [emailUser, mobileUser] = await Promise.all([
    User.findOne({ email, _id: { $ne: req.auth.sub } }).select("_id"),
    User.findOne({ country, mobile, _id: { $ne: req.auth.sub } }).select("_id")
  ]);

  if (emailUser) {
    return res.status(409).json({ message: "Another account already uses this email address." });
  }

  if (mobileUser) {
    return res.status(409).json({ message: "Another account already uses this country and mobile number." });
  }

  await User.updateOne(
    { _id: req.auth.sub, role: ROLE_ASSOCIATE_MEMBER },
    {
      $set: {
        ownerName,
        businessName,
        email,
        country,
        mobile,
        state,
        district,
        city,
        pinCode,
        address,
        lastActivityAt: new Date()
      }
    }
  );

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Associate Member Profile",
    action: "Update Profile",
    targetType: "Associate Member",
    targetId: req.auth.sub,
    targetName: actorName,
    details: {
      email,
      mobile
    }
  });

  return res.json({
    message: "Profile updated successfully."
  });
});

app.patch("/api/associate-member-module/change-password", authenticate, requireRole(ROLE_ASSOCIATE_MEMBER), ensureDatabaseConnected, async (req, res) => {
  const associateMember = await User.findOne({ _id: req.auth.sub, role: ROLE_ASSOCIATE_MEMBER }).select("_id ownerName passwordHash");
  if (!associateMember) {
    return res.status(404).json({ message: "Associate Member record not found." });
  }

  const currentPassword = String(req.body.currentPassword || "");
  const newPassword = String(req.body.newPassword || "");
  const confirmPassword = String(req.body.confirmPassword || "");

  if (!currentPassword || !newPassword || !confirmPassword) {
    return res.status(400).json({ message: "Current password, new password, and confirm password are required." });
  }

  const isCurrentPasswordValid = await bcrypt.compare(currentPassword, associateMember.passwordHash);
  if (!isCurrentPasswordValid) {
    return res.status(400).json({ message: "Current password is incorrect." });
  }

  if (newPassword.length < 8) {
    return res.status(400).json({ message: "New password must be at least 8 characters long." });
  }

  if (newPassword !== confirmPassword) {
    return res.status(400).json({ message: "New password and confirm password do not match." });
  }

  if (currentPassword === newPassword) {
    return res.status(400).json({ message: "New password must be different from the current password." });
  }

  const passwordHash = await bcrypt.hash(newPassword, 12);

  await User.updateOne(
    { _id: req.auth.sub, role: ROLE_ASSOCIATE_MEMBER },
    {
      $set: {
        passwordHash,
        lastActivityAt: new Date()
      }
    }
  );

  const actorName = await getActorName(req.auth);
  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Associate Member Profile",
    action: "Change Password",
    targetType: "Associate Member",
    targetId: req.auth.sub,
    targetName: associateMember.ownerName || actorName,
    details: {
      message: "Associate member changed password successfully."
    }
  });

  return res.json({ message: "Password updated successfully." });
});

async function buildAdminModuleBootstrap(adminId) {
  const assignedMemberIds = await User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId }).distinct("_id");
  const [
    adminUser,
    totalMembers,
    totalOrders,
    pendingOrders,
    printingOrders,
    packagingOrders,
    dispatchOrders,
    completedOrders,
    pendingTopUps,
    totalTopUpTransactions,
    totalWalletDebitTransactions,
    walletBalance,
    recentUsers,
    recentOrders,
    recentWalletRequests,
    recentWalletDebitOrders,
    pendingMembers,
    pendingOrderItems,
    pendingWalletRequests
  ] = await Promise.all([
    User.findOne({ _id: adminId, role: ROLE_ADMIN }).select("ownerName mobile email businessName status role updatedAt createdAt"),
    User.countDocuments({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId }),
    Order.countDocuments({ assignedAdminId: adminId }),
    Order.countDocuments({ assignedAdminId: adminId, status: ORDER_STATUS_PENDING }),
    Order.countDocuments({ assignedAdminId: adminId, status: ORDER_STATUS_PRINTING }),
    Order.countDocuments({ assignedAdminId: adminId, status: ORDER_STATUS_PACKAGING }),
    Order.countDocuments({ assignedAdminId: adminId, status: ORDER_STATUS_DISPATCHED }),
    Order.countDocuments({ assignedAdminId: adminId, status: ORDER_STATUS_COMPLETED }),
    TopUpRequest.countDocuments({ requestedByUserId: { $in: assignedMemberIds }, status: "pending" }),
    TopUpRequest.countDocuments({ requestedByUserId: { $in: assignedMemberIds } }),
    Order.countDocuments({ assignedAdminId: adminId, walletDebitAmount: { $gt: 0 } }),
    getWalletBalanceForActor({ userId: adminId, role: ROLE_ADMIN }),
    User.find({
      $or: [{ _id: adminId, role: ROLE_ADMIN }, { role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId }]
    })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .select("ownerName businessName mobile status role updatedAt createdAt"),
    Order.find({ assignedAdminId: adminId })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("placedByUserId", "ownerName businessName")
      .select("orderNumber orderName customerName status updatedAt createdAt placedByUserId"),
    TopUpRequest.find({ requestedByUserId: { $in: assignedMemberIds } })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("requestedByUserId", "ownerName businessName mobile role")
      .select("requestedByUserId requestedByDisplayName amount status updatedAt createdAt"),
    Order.find({ assignedAdminId: adminId, walletDebitAmount: { $gt: 0 } })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("placedByUserId", "ownerName businessName mobile role")
      .select("orderNumber orderName customerName status updatedAt createdAt walletDebitAmount referenceNo placedByUserId"),
    User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId, status: STATUS_PENDING })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .select("ownerName businessName mobile status updatedAt createdAt"),
    Order.find({ assignedAdminId: adminId, status: ORDER_STATUS_PENDING })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .select("orderNumber orderName customerName status updatedAt createdAt"),
    TopUpRequest.find({ requestedByUserId: { $in: assignedMemberIds }, status: "pending" })
      .sort({ updatedAt: -1, createdAt: -1 })
      .limit(6)
      .populate("requestedByUserId", "ownerName businessName mobile role")
      .select("requestedByUserId requestedByDisplayName amount status updatedAt createdAt")
  ]);

  const totalUsers = totalMembers + 1;
  const totalWalletTransactions = totalTopUpTransactions + totalWalletDebitTransactions;

  const recentUserActivities = recentUsers.map((user) => ({
    id: user._id.toString(),
    name: user.ownerName || "--",
    role: normalizeRoleLabel(user.role || ""),
    businessName: user.businessName || "--",
    mobile: user.mobile || "--",
    status: user.role === ROLE_ADMIN ? normalizeAdminStatusLabel(user.status) : formatSystemStatusLabel(user.status),
    statusValue: String(user.status || "").trim().toLowerCase(),
    updatedOn: formatDateTime(user.updatedAt || user.createdAt),
    detailsPath:
      user.role === ROLE_ADMIN
        ? null
        : `/dashboard/admin/associate-members/details/${user._id.toString()}`
  }));

  const recentOrderActivities = recentOrders.map((order) => ({
    id: order._id.toString(),
    orderNumber: order.orderNumber ? `#${order.orderNumber}` : "--",
    orderName: order.orderName || "--",
    createdBy: order.placedByUserId?.ownerName || order.customerName || "--",
    status: normalizeOrderStatusLabel(order.status),
    statusValue: String(order.status || "").trim().toLowerCase(),
    updatedOn: formatDateTime(order.updatedAt || order.createdAt),
    detailsPath: `/dashboard/admin/orders/details/${order._id.toString()}`
  }));

  const recentWalletCredits = recentWalletRequests.map((request) => ({
    id: request._id.toString(),
    reference: `TPR-${request._id.toString().slice(-6).toUpperCase()}`,
    type: "Credit",
    memberName: request.requestedByUserId?.ownerName || request.requestedByDisplayName || "--",
    amount: Number(request.amount || 0).toFixed(2),
    status: normalizeTopUpStatusLabel(request.status),
    statusValue: String(request.status || "").trim().toLowerCase(),
    businessName: request.requestedByUserId?.businessName || "--",
    updatedOn: formatDateTime(request.updatedAt || request.createdAt),
    sortDate: request.updatedAt || request.createdAt || new Date(0),
    detailsPath: `/dashboard/admin/wallet/details/${request._id.toString()}`
  }));

  const recentWalletDebits = recentWalletDebitOrders.map((order) => ({
    id: `wallet-debit-${order._id.toString()}`,
    reference: order.referenceNo || (order.orderNumber ? `ORD-${order.orderNumber}` : `ORD-${order._id.toString().slice(-6).toUpperCase()}`),
    type: "Debit",
    memberName: order.placedByUserId?.ownerName || order.customerName || "--",
    amount: Number(order.walletDebitAmount || 0).toFixed(2),
    status: normalizeOrderStatusLabel(order.status),
    statusValue: String(order.status || "").trim().toLowerCase(),
    businessName: order.placedByUserId?.businessName || "--",
    updatedOn: formatDateTime(order.updatedAt || order.createdAt),
    sortDate: order.updatedAt || order.createdAt || new Date(0),
    detailsPath: `/dashboard/admin/orders/details/${order._id.toString()}`
  }));

  const recentWalletTransactions = [...recentWalletCredits, ...recentWalletDebits]
    .sort((left, right) => new Date(right.sortDate).getTime() - new Date(left.sortDate).getTime())
    .slice(0, 6)
    .map(({ sortDate, ...transaction }) => transaction);

  const pendingTasks = [
    ...pendingMembers.map((member) => ({
      id: `member-${member._id.toString()}`,
      module: "User Management",
      title: member.ownerName || "Associate Member",
      status: formatSystemStatusLabel(member.status),
      statusValue: String(member.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(member.updatedAt || member.createdAt),
      actionLabel: "Review",
      path: `/dashboard/admin/associate-members/details/${member._id.toString()}`,
      sortDate: member.updatedAt || member.createdAt || new Date(0)
    })),
    ...pendingOrderItems.map((order) => ({
      id: `order-${order._id.toString()}`,
      module: "Order Management",
      title: order.orderNumber ? `Order #${order.orderNumber}` : order.orderName || "--",
      status: normalizeOrderStatusLabel(order.status),
      statusValue: String(order.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(order.updatedAt || order.createdAt),
      actionLabel: "Process",
      path: `/dashboard/admin/orders/details/${order._id.toString()}`,
      sortDate: order.updatedAt || order.createdAt || new Date(0)
    })),
    ...pendingWalletRequests.map((request) => ({
      id: `wallet-${request._id.toString()}`,
      module: "Wallet Management",
      title: request.requestedByUserId?.ownerName || request.requestedByDisplayName || "Wallet Request",
      status: normalizeTopUpStatusLabel(request.status),
      statusValue: String(request.status || "").trim().toLowerCase(),
      updatedOn: formatDateTime(request.updatedAt || request.createdAt),
      actionLabel: "Review",
      path: `/dashboard/admin/wallet/details/${request._id.toString()}`,
      sortDate: request.updatedAt || request.createdAt || new Date(0)
    }))
  ]
    .sort((left, right) => new Date(right.sortDate).getTime() - new Date(left.sortDate).getTime())
    .slice(0, 10)
    .map(({ sortDate, ...task }) => task);

  return {
    header: {
      systemName: "Printing Services Division",
      systemDescription: "Header, navigation bar, and footer preserve the attached admin design language.",
      memberId: `ADM-${String(adminUser?.mobile || "0000").slice(-4).padStart(4, "0")}`,
      accountBalance: walletBalance.toFixed(2)
    },
    footer: {
      companyName: "Printers Club of India Limited",
      companyDescription:
        "Dedicated to the continuous development and modernization of the printing industry in India. Providing quality services and a unified platform for printers nationwide.",
      contactEmail: "support@printersclub.in",
      contactPhone: "+91 99758 13249",
      contactAddress: "Plot No. 57, Jhotwara Industrial Area, Jaipur, Rajasthan",
      quickLinks: [
        { label: "Dashboard", href: "/dashboard/admin" },
        { label: "Associate Members", href: "/dashboard/admin/associate-members/all" },
        { label: "Orders", href: "/dashboard/admin/orders/all" },
        { label: "Wallet", href: "/dashboard/admin/wallet/add-money" }
      ],
      socialLinks: [
        { label: "Facebook", href: "#" },
        { label: "Instagram", href: "#" },
        { label: "LinkedIn", href: "#" }
      ],
      termsLink: "Terms_And_Conditions.aspx?type=new"
    },
    dashboardCards: [
      { title: "Total Users", value: totalUsers, note: "Your admin account plus assigned associate members.", iconKey: "associate-members" },
      { title: "Total Associate Members", value: totalMembers, note: "Associate members currently mapped to this admin.", iconKey: "associate-members" },
      { title: "Total Orders", value: totalOrders, note: "Orders currently within this admin scope.", iconKey: "orders" },
      { title: "Pending Orders", value: pendingOrders, note: "Orders waiting for review or production.", iconKey: "orders" },
      { title: "Printing Orders", value: printingOrders, note: "Orders currently in printing.", iconKey: "orders" },
      { title: "Packaging Orders", value: packagingOrders, note: "Orders currently in packaging.", iconKey: "orders" },
      { title: "Dispatch Orders", value: dispatchOrders, note: "Orders currently in dispatch.", iconKey: "orders" },
      { title: "Completed Orders", value: completedOrders, note: "Orders completed under this admin.", iconKey: "orders" },
      { title: "Total Wallet Transactions", value: totalWalletTransactions, note: "Top-up credits and order debits in your scope.", iconKey: "wallet" },
      { title: "Pending Wallet Requests", value: pendingTopUps, note: "Wallet requests waiting for review.", iconKey: "wallet" }
    ],
    dashboardOverview: {
      totalUsers,
      totalAssociateMembers: totalMembers,
      totalOrders,
      pendingOrders,
      printingOrders,
      packagingOrders,
      dispatchOrders,
      completedOrders,
      totalWalletTransactions,
      pendingWalletRequests: pendingTopUps,
      accountBalance: walletBalance.toFixed(2)
    },
    recentActivities: {
      users: recentUserActivities,
      orders: recentOrderActivities,
      walletTransactions: recentWalletTransactions
    },
    pendingTasks,
    recentTasks: pendingTasks,
    moduleCards: [
      { title: "Associate Member Management", description: "All associate-member routes match the requested menu tree and are wired for API-backed tables.", meta: `${totalMembers} available records` },
      { title: "Order Management", description: "Status-specific order routes are ready for pending, printing, packaging, dispatch, completed, improper, cancelled, and rejected views.", meta: `${totalOrders} tracked orders` },
      { title: "Wallet Management", description: "Wallet transactions and top-up request states are grouped under the same admin navigation shell.", meta: `${pendingTopUps} pending actions` },
      { title: "Reports & Activity Logs", description: "Report, notification, activity log, and profile routes are all exposed through reusable React components.", meta: `${pendingTasks.length} pending task records` },
      { title: "MERN Ready Structure", description: "This module now uses React, Tailwind, Context API, Axios, Express, MongoDB, JWT auth, and RBAC-safe endpoints.", meta: "Reusable shell components" }
    ]
  };
}

async function buildAdminModuleSectionData(section, view, adminId) {
  if (section === "associate-members") {
    const query = { role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId };

    if (view === "active") {
      query.status = STATUS_ACTIVE;
    } else if (view === "inactive") {
      query.status = STATUS_SUSPENDED;
    }

    const members = await User.find(query)
      .sort({ updatedAt: -1 })
      .limit(50)
      .select("ownerName mobile email businessName city district state status updatedAt");

    return {
      meta: {
        columns: [
          { key: "name", label: "Member Name" },
          { key: "mobile", label: "WhatsApp" },
          { key: "email", label: "Email" },
          { key: "businessName", label: "Business / Firm" },
          { key: "status", label: "Status" },
          { key: "location", label: "Location" }
        ]
      },
      summary: {
        total: members.length,
        active: members.filter((member) => member.status === STATUS_ACTIVE).length,
        inactive: members.filter((member) => member.status === STATUS_SUSPENDED).length
      },
      items: members.map((member) => ({
        id: member._id.toString(),
        name: member.ownerName,
        mobile: member.mobile,
        email: member.email,
        businessName: member.businessName,
        status: normalizeAdminStatusLabel(member.status),
        location: [member.city, member.district, member.state].filter(Boolean).join(", ")
      }))
    };
  }

  if (section === "orders" || section === "production") {
    const query = { assignedAdminId: adminId };
    const allowedStatuses = section === "production" ? getAdminModuleOrderStatuses(view) : getAdminModuleOrderStatuses(view);

    if (allowedStatuses.length > 0) {
      query.status = { $in: allowedStatuses };
    }

    const orders = await Order.find(query)
      .sort({ updatedAt: -1 })
      .limit(50)
      .populate("placedByUserId", "ownerName businessName mobile email")
      .select(
        "orderNumber orderName customerName customerMobile status designSubmissionSource designFileType designFileName orderDetailsOverview orderedAt createdAt placedByUserId"
      );

    return {
      meta: {
        columns: [
          { key: "orderNumber", label: "Order No." },
          { key: "dateTime", label: "Date" },
          { key: "orderName", label: "Order Name" },
          { key: "createdBy", label: "Created By" },
          { key: "orderDetail", label: "Order Detail" },
          { key: "status", label: "Status" },
          { key: "fileType", label: "File Type" }
        ]
      },
      summary: {
        total: orders.length,
        pending: orders.filter((order) => order.status === ORDER_STATUS_PENDING).length,
        completed: orders.filter((order) => order.status === ORDER_STATUS_COMPLETED).length
      },
      items: orders.map((order) => ({
        id: order._id.toString(),
        orderNumber: order.orderNumber || "--",
        dateTime: formatDateTime(order.orderedAt || order.createdAt || order.updatedAt),
        orderName: order.orderName || "--",
        createdBy: order.placedByUserId?.ownerName || order.customerName || "--",
        orderDetail: order.orderDetailsOverview || "--",
        status: normalizeOrderStatusLabel(order.status),
        fileSource: normalizeDesignSubmissionSourceLabel(order.designSubmissionSource) || "--",
        fileType:
          String(order.designSubmissionSource || "").toLowerCase() === DESIGN_SOURCE_EMAIL
            ? "Email"
            : String(order.designFileType || (order.designFileName?.includes(".") ? order.designFileName.split(".").pop() : "File"))
                .trim()
                .toUpperCase() || "File",
        hasEmailDesign: String(order.designSubmissionSource || "").toLowerCase() === DESIGN_SOURCE_EMAIL
      }))
    };
  }

  if (section === "wallet") {
    const assignedMemberIds = await User.find({ role: ROLE_ASSOCIATE_MEMBER, assignedAdminId: adminId }).distinct("_id");
    const walletQuery = { requestedByUserId: { $in: assignedMemberIds } };

    if (view === "pending") walletQuery.status = "pending";
    if (view === "approved") walletQuery.status = "approved";
    if (view === "rejected") walletQuery.status = "rejected";

    const requests = await TopUpRequest.find(walletQuery)
      .sort({ createdAt: -1 })
      .limit(50)
      .populate("requestedByUserId", "ownerName mobile businessName");

    return {
      meta: {
        columns: [
          { key: "reference", label: "Reference" },
          { key: "memberName", label: "Member Name" },
          { key: "mobile", label: "Mobile" },
          { key: "amount", label: "Amount" },
          { key: "status", label: "Status" },
          { key: "businessName", label: "Business / Firm" }
        ]
      },
      summary: {
        total: requests.length,
        pending: requests.filter((request) => request.status === "pending").length,
        approved: requests.filter((request) => request.status === "approved").length,
        rejected: requests.filter((request) => request.status === "rejected").length
      },
      items: requests.map((request) => ({
        id: request._id.toString(),
        reference: `TPR-${request._id.toString().slice(-6).toUpperCase()}`,
        memberName: request.requestedByUserId?.ownerName || "--",
        mobile: request.requestedByUserId?.mobile || "--",
        amount: Number(request.amount || 0).toFixed(2),
        status: String(request.status || "").replace(/^\w/, (char) => char.toUpperCase()),
        businessName: request.requestedByUserId?.businessName || "--"
      }))
    };
  }

  if (section === "activity-logs") {
    const logs = await AuditLog.find({ $or: [{ actorId: adminId }, { "details.assignedAdminId": String(adminId) }] })
      .sort({ createdAt: -1 })
      .limit(50);

    return {
      meta: {
        columns: [
          { key: "module", label: "Module" },
          { key: "action", label: "Action" },
          { key: "actor", label: "User" },
          { key: "target", label: "Target" },
          { key: "status", label: "Status" },
          { key: "date", label: "Date & Time" }
        ]
      },
      summary: {
        total: logs.length
      },
      items: logs.map((log) => ({
        id: log._id.toString(),
        module: log.module,
        action: log.action,
        actor: log.actorName,
        target: log.targetName || "--",
        status: log.details?.nextStatus || "Updated",
        date: formatDateTime(log.createdAt)
      }))
    };
  }

  if (section === "notifications") {
    const orders = await Order.find({ assignedAdminId: adminId }).sort({ updatedAt: -1 }).limit(20).select("orderName orderNumber status updatedAt");

    return {
      meta: {
        columns: [
          { key: "title", label: "Notification" },
          { key: "message", label: "Message" },
          { key: "type", label: "Type" },
          { key: "status", label: "Status" },
          { key: "date", label: "Updated At" }
        ]
      },
      summary: {
        total: orders.length
      },
      items: orders.map((order) => ({
        id: order._id.toString(),
        title: `Order ${order.orderNumber || "--"}`,
        message: `${order.orderName || "Order"} is currently ${normalizeOrderStatusLabel(order.status)}.`,
        type: "Order Update",
        status: "Unread",
        date: formatDateTime(order.updatedAt)
      }))
    };
  }

  if (section === "profile") {
    const admin = await User.findOne({ _id: adminId, role: ROLE_ADMIN }).select(
      "ownerName mobile email businessName country state district city pinCode address"
    );

    return {
      meta: {
        columns: [
          { key: "field", label: "Field" },
          { key: "value", label: "Value" }
        ]
      },
      summary: {
        total: admin ? 8 : 0
      },
      items: admin
        ? [
            { id: "name", field: "Admin Name", value: admin.ownerName },
            { id: "business", field: "Business / Firm", value: admin.businessName },
            { id: "mobile", field: "WhatsApp Number", value: admin.mobile },
            { id: "email", field: "Email Address", value: admin.email },
            { id: "country", field: "Country", value: admin.country },
            { id: "state", field: "State", value: admin.state },
            { id: "district", field: "District", value: admin.district },
            { id: "address", field: "Full Address", value: buildAddress(admin) }
          ]
        : []
    };
  }

  return {
    meta: {
      columns: [
        { key: "title", label: "Module" },
        { key: "description", label: "Description" },
        { key: "status", label: "Status" }
      ]
    },
    summary: {
      total: 4
    },
    items: [
      { id: "orders-report", title: "Order Reports", description: "Export orders in Excel, PDF, and CSV.", status: "Ready" },
      { id: "associate-report", title: "Associate Member Reports", description: "Generate member-level operational summaries.", status: "Ready" },
      { id: "wallet-report", title: "Wallet Reports", description: "Track wallet requests and transactions.", status: "Ready" },
      { id: "production-report", title: "Production Reports", description: "Review printing, packaging, and dispatch queue data.", status: "Ready" }
    ]
  };
}

app.get("/api/admin-module/bootstrap", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return res.json(await buildAdminModuleBootstrap(req.auth.sub));
});

app.get("/api/admin-module/orders/:orderId", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Order record ID." });
  }

  const order = await Order.findOne({ _id: req.params.orderId, assignedAdminId: req.auth.sub });
  if (!order) {
    return res.status(404).json({ message: "Order record not found." });
  }

  return res.json({ order: await buildOrderDetails(order) });
});

app.get("/api/admin-module/associate-members/:associateMemberId", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.associateMemberId)) {
    return res.status(400).json({ message: "Invalid Associate Member record ID." });
  }

  const associateMember = await User.findOne({
    _id: req.params.associateMemberId,
    role: ROLE_ASSOCIATE_MEMBER,
    assignedAdminId: req.auth.sub
  });
  if (!associateMember) {
    return res.status(404).json({ message: "Associate Member record not found." });
  }

  return res.json({
    associateMember: await buildAssociateMemberDetails(associateMember)
  });
});

app.patch("/api/admin-module/orders/:orderId/status", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  if (!mongoose.Types.ObjectId.isValid(req.params.orderId)) {
    return res.status(400).json({ message: "Invalid Order record ID." });
  }

  const allowedStatuses = [
    ORDER_STATUS_PENDING,
    ORDER_STATUS_PRINTING,
    ORDER_STATUS_PACKAGING,
    ORDER_STATUS_DISPATCHED,
    ORDER_STATUS_COMPLETED,
    ORDER_STATUS_IMPROPER,
    ORDER_STATUS_CANCELLED
  ];
  const nextStatus = normalizeOrderStatusInput(req.body.status);
  if (!allowedStatuses.includes(nextStatus)) {
    return res.status(400).json({
      message: "Order status must be Pending, Printing, Packaging, Dispatched, Completed, Improper, or Cancelled."
    });
  }

  const actorName = await getActorName(req.auth);
  const note = String(req.body.note || "").trim();

  const order = await Order.findOneAndUpdate(
    { _id: req.params.orderId, assignedAdminId: req.auth.sub },
    {
      $set: {
        status: nextStatus,
        updatedAt: new Date()
      },
      $push: {
        statusHistory: {
          status: nextStatus,
          note,
          changedAt: new Date(),
          changedByRole: req.auth.role,
          changedById: req.auth.sub,
          changedByName: actorName
        }
      }
    },
    { new: true }
  );

  if (!order) {
    return res.status(404).json({ message: "Order record not found." });
  }

  await createAuditLog({
    actorRole: req.auth.role,
    actorId: req.auth.sub,
    actorName,
    module: "Admin Orders Management",
    action: `Update Order Status to ${normalizeOrderStatusLabel(nextStatus)}`,
    targetType: "Order",
    targetId: order._id.toString(),
    targetName: order.orderName || `Order ${order.orderNumber}`,
    details: {
      orderNumber: order.orderNumber,
      nextStatus: normalizeOrderStatusLabel(nextStatus),
      note
    }
  });

  return res.json({
    message: `Order status updated to ${normalizeOrderStatusLabel(nextStatus)} successfully.`,
    order: await buildOrderDetails(order)
  });
});

app.get("/api/admin-module/section/:section", authenticate, requireRole(ROLE_ADMIN), ensureDatabaseConnected, async (req, res) => {
  return res.json(await buildAdminModuleSectionData(req.params.section, req.query.view, req.auth.sub));
});

app.use("/api", (_req, res) => {
  return res.status(404).json({
    message: "API endpoint not found. If you recently added this route, restart the backend server and try again."
  });
});

async function startServer() {
  if (process.env.MONGO_URI) {
    await mongoose.connect(process.env.MONGO_URI);
    console.log("MongoDB connected");
  }

  const server = app.listen(port, () => {
    console.log(`API running on http://localhost:${port}`);
  });

  server.on("error", (error) => {
    if (error.code === "EADDRINUSE") {
      console.error(`Port ${port} is already in use. Stop the running process or change PORT in backend/.env.`);
      process.exit(1);
    }

    throw error;
  });
}

startServer().catch((error) => {
  console.error("Server failed to start", error);
  process.exit(1);
});
