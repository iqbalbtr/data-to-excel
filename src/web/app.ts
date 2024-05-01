import express from "express";
import { router } from "../routes/api";

export const app = express();

app.use(express.json());
app.use(express.static("public"))

app.use(router);