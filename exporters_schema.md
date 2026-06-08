# Exporters Database Schema

Generated: 2026-06-08T15:30:16+00:00

## Overview
This project stores exporter listings, contacts, certifications, accreditations, products, and reporting views in exporters_final.db.

## Main Reporting Views
- **v_company_profile**: one row per company with certifications and accreditations rolled up.
- **v_company_products**: one row per company with product family and product variant rollups.
- **v_product_hierarchy**: product family hierarchy reference.

## Known Limitations
- Product matching is keyword based, not ML-based.
- Contacts are extracted from visible text and links, so some false positives/negatives can still occur.
- Some exporter pages contain repeated site boilerplate which is filtered on a best-effort basis.

## Index: idx_companies_exporter_type

```sql
CREATE INDEX idx_companies_exporter_type ON companies(exporter_type);
```

## Index: idx_companies_name

```sql
CREATE INDEX idx_companies_name ON companies(name);
```

## Index: idx_companies_state

```sql
CREATE INDEX idx_companies_state ON companies(state);
```

## Index: idx_company_products_company

```sql
CREATE INDEX idx_company_products_company ON company_products(company_id);
```

## Table: accreditations

Accreditation lookup table.

```sql
CREATE TABLE accreditations (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE);
```

## Table: attributes

Raw extracted text attributes kept for traceability.

```sql
CREATE TABLE attributes (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, name TEXT, value TEXT,FOREIGN KEY(company_id) REFERENCES companies(id));
```

## Table: certifications

Certification lookup table.

```sql
CREATE TABLE certifications (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE);
```

## Table: companies

Master exporter/company records.

```sql
CREATE TABLE companies (id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT, description TEXT, profile_url TEXT UNIQUE, image_url TEXT,exporter_type TEXT, licence_number TEXT, establishment_numbers TEXT,website TEXT, address TEXT, abn TEXT, state TEXT, postcode TEXT,page_title TEXT, page_heading TEXT, meta_description TEXT, meta_title TEXT,page_text_excerpt TEXT, details_json TEXT, profile_error TEXT);
```

## Table: company_accreditations

Many-to-many link between companies and accreditations.

```sql
CREATE TABLE company_accreditations (company_id INTEGER, accreditation_id INTEGER, UNIQUE(company_id, accreditation_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(accreditation_id) REFERENCES accreditations(id));
```

## Table: company_certifications

Many-to-many link between companies and certifications.

```sql
CREATE TABLE company_certifications (company_id INTEGER, certification_id INTEGER, UNIQUE(company_id, certification_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(certification_id) REFERENCES certifications(id));
```

## Table: company_products

Matched company-to-product relationships.

```sql
CREATE TABLE company_products (company_id INTEGER, product_id INTEGER, match_source TEXT, source_field TEXT, matched_keyword TEXT, match_confidence REAL,UNIQUE(company_id, product_id, matched_keyword, source_field),FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(product_id) REFERENCES products(id));
```

## Table: contacts

Extracted contact details by company.

```sql
CREATE TABLE contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, email TEXT, phone TEXT, contact_source TEXT,FOREIGN KEY(company_id) REFERENCES companies(id));
```

## Table: product_families

Top-level product family dimension.

```sql
CREATE TABLE product_families (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, sort_order INTEGER);
```

## Table: product_family_matches

Matched company-to-product-family relationships.

```sql
CREATE TABLE product_family_matches (company_id INTEGER, family_id INTEGER, UNIQUE(company_id, family_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(family_id) REFERENCES product_families(id));
```

## Table: product_match_audit

Audit table showing why product matches were made.

```sql
CREATE TABLE product_match_audit (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, product_id INTEGER, family_id INTEGER, source_field TEXT, matched_keyword TEXT, source_excerpt TEXT, match_confidence REAL,FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(product_id) REFERENCES products(id), FOREIGN KEY(family_id) REFERENCES product_families(id));
```

## Table: products

Product dimension with hierarchy metadata.

```sql
CREATE TABLE products (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, family_id INTEGER, parent_product_id INTEGER, hierarchy_level INTEGER, FOREIGN KEY(family_id) REFERENCES product_families(id), FOREIGN KEY(parent_product_id) REFERENCES products(id));
```

## View: v_company_products

Company-product reporting view.

```sql
CREATE VIEW v_company_products AS SELECT c.id, c.name, c.exporter_type, c.website, c.state, c.profile_url, GROUP_CONCAT(DISTINCT pf.name) AS product_families, GROUP_CONCAT(DISTINCT CASE WHEN p.hierarchy_level = 2 THEN p.name END) AS product_variants FROM companies c LEFT JOIN company_products cp ON c.id = cp.company_id LEFT JOIN products p ON cp.product_id = p.id LEFT JOIN product_families pf ON p.family_id = pf.id GROUP BY c.id;
```

## View: v_company_profile

Profile reporting view.

```sql
CREATE VIEW v_company_profile AS SELECT c.id, c.name, c.exporter_type, c.website, c.address, c.state, c.postcode, c.profile_url, GROUP_CONCAT(DISTINCT cert.name) AS certifications, GROUP_CONCAT(DISTINCT acc.name) AS accreditations FROM companies c LEFT JOIN company_certifications cc ON c.id = cc.company_id LEFT JOIN certifications cert ON cc.certification_id = cert.id LEFT JOIN company_accreditations ca ON c.id = ca.company_id LEFT JOIN accreditations acc ON ca.accreditation_id = acc.id GROUP BY c.id;
```

## View: v_product_hierarchy

Hierarchy reference view.

```sql
CREATE VIEW v_product_hierarchy AS SELECT pf.name AS family_name, root.name AS root_product, child.name AS variant_product, child.hierarchy_level FROM product_families pf LEFT JOIN products root ON root.family_id = pf.id AND root.parent_product_id IS NULL LEFT JOIN products child ON child.parent_product_id = root.id ORDER BY pf.sort_order, child.name;
```
