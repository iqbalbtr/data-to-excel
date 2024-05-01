import express from "express";
import webController from "../controller/web-controller";

export const router = express.Router();

router.get("/", webController.get);
router.get("/excel", webController.excel);