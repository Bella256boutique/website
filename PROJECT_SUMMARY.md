# Bella256 Boutique — Ecommerce Build (Project Context)

## Owner
Lisa Lyles — Certified Mastectomy Fitter, Madison, GA. Bella256 Boutique LLC serves mastectomy and medical hair-loss patients. Little website-building experience: explain every step, one at a time, bullet points, no fluff. Messaging rule: lead with "certified mastectomy fitter" — do NOT lead with RN/nurse.

## Stack
- WordPress on Bluehost, bella256boutique.com
- Theme: Kadence + Kadence Blocks
- WooCommerce: plugin installed (Sept 2026), setup wizard NOT yet completed
- Payments: PayPal exists; Stripe planned (no account yet — set up at payments step)
- Store is self-pay only (not Medicare-enrolled; no insurance billing)

## Branding (extracted from live site — use exactly)
- Dusty rose #D3A4A5 (primary), darker rose #BC8A8B
- Soft teal #8FC5C5 (accents/announcement bar)
- Warm blush #E9D6D1, off-white #F7F0EF, charcoal #3F3F3F
- Fonts in mockups: Poppins (headings), Inter (body)
- Tagline: "You don't have to face breast cancer alone."

## Site architecture (approved)
- Need-centered, NOT fittings-centered. Homepage hero: "What Do You Need Today?" with 3 buttons: Insurance Reimbursement | Shop Self-Pay | Schedule a Fitting
- Shop by Need categories: Mastectomy Bras, Breast Forms, Recovery, Compression, Medical Wigs (+ Accessories)
- "Bella256 Fit & Care Quiz" (5 questions, conditional routing) — phase 2
- Insurance framing = SUPERBILL model: customer shops self-pay, receives superbill (NPI + HCPCS L8000/L8020/L8030 etc.) for out-of-network commercial reimbursement. Medicare canNOT be reimbursed (non-enrolled supplier) — say so honestly. Never gate the shop behind insurance verification.
- Inspirations: Myya (clinical funnel), AnaOno (survivor UX), MastectomyShop (SEO categories)

## Vendors & products
- ABC (American Breast Care), acct 7718: 99 products (Bras 32, Breast Forms 37, Shapers 12, Accessories 19 incl. 1 blank row). Retail = ABC List Price (her cost is 8% below list — her explicit choice). 10 accessories have no price → imported as drafts (styles 941, 927, 928, 920, 951P, 924, 921, 903, 922, 901).
- Jon Renau: dropship wigs, retailer account approved, catalog NOT built yet (no public feed/API — ask account manager for data files). Feeds the "Medical Wigs" category.

## Deliverables already produced (in earlier Cowork session)
- bella256_woocommerce_import.csv — 181 rows: 99 products (variable products with Color/Size attributes; one "Any/Any" variation each carries the price), need-based categories, ABC image URLs
- Website architecture blueprint (docx): sitemap, homepage copy, SEO slugs/meta titles, quiz logic, 10-step build checklist
- Homepage + Mastectomy Bras shop page HTML mockups (approved branding)
- Superbill template (docx): supplier/patient/prescriber blocks, ICD-10 Z90.11-.13, HCPCS reference, page-2 patient submission guide

## Build roadmap (current position: step 1 done except wizard)
1. Install WooCommerce ✅ (wizard pending: Madison GA address, physical products, skip all extras, payments later)
2. Store settings — address, GA sales tax, currency
3. Import 99 ABC products via Products → Import (CSV above)
4. Kadence global colors = brand palette; build homepage sections per mockup
5. Shipping zones & policies
6. Payments — PayPal, then Stripe
7. End-to-end test order before launch
