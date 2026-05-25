import { z } from "zod";

// No database needed - this is a stateless document generator
// Define API request/response schemas for type safety

export const googleSheetRequestSchema = z.object({
  url: z.string().url(),
});

export const aiRewriteRequestSchema = z.object({
  text: z.string().min(1),
});

export const aiRewriteResponseSchema = z.object({
  originalText: z.string(),
  rewrittenText: z.string(),
});

export const proxyImageRequestSchema = z.object({
  url: z.string().url(),
});

export type GoogleSheetRequest = z.infer<typeof googleSheetRequestSchema>;
export type AIRewriteRequest = z.infer<typeof aiRewriteRequestSchema>;
export type AIRewriteResponse = z.infer<typeof aiRewriteResponseSchema>;
export type ProxyImageRequest = z.infer<typeof proxyImageRequestSchema>;
