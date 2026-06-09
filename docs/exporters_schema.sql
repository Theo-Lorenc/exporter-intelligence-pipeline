CREATE INDEX idx_companies_exporter_type ON companies(exporter_type);

CREATE INDEX idx_companies_name ON companies(name);

CREATE INDEX idx_companies_state ON companies(state);

CREATE INDEX idx_company_products_company ON company_products(company_id);

CREATE INDEX idx_contacts_company ON contacts(company_id);

CREATE TABLE accreditations (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE);

CREATE TABLE attributes (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, name TEXT, value TEXT,FOREIGN KEY(company_id) REFERENCES companies(id));

CREATE TABLE certifications (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE);

CREATE TABLE companies (id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT, description TEXT, profile_url TEXT UNIQUE, image_url TEXT,exporter_type TEXT, licence_number TEXT, establishment_numbers TEXT,website TEXT, address TEXT, abn TEXT, state TEXT, postcode TEXT,countries_served TEXT,page_title TEXT, page_heading TEXT, meta_description TEXT, meta_title TEXT,page_text_excerpt TEXT, details_json TEXT, profile_error TEXT);

CREATE TABLE company_accreditations (company_id INTEGER, accreditation_id INTEGER, UNIQUE(company_id, accreditation_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(accreditation_id) REFERENCES accreditations(id));

CREATE TABLE company_certifications (company_id INTEGER, certification_id INTEGER, UNIQUE(company_id, certification_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(certification_id) REFERENCES certifications(id));

CREATE TABLE company_products (company_id INTEGER, product_id INTEGER, match_source TEXT, source_field TEXT, matched_keyword TEXT, match_confidence REAL,UNIQUE(company_id, product_id, matched_keyword, source_field),FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(product_id) REFERENCES products(id));

CREATE TABLE contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, email TEXT, phone TEXT, contact_source TEXT,UNIQUE(company_id, email, phone, contact_source),FOREIGN KEY(company_id) REFERENCES companies(id));

CREATE TABLE product_families (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, sort_order INTEGER);

CREATE TABLE product_family_matches (company_id INTEGER, family_id INTEGER, UNIQUE(company_id, family_id), FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(family_id) REFERENCES product_families(id));

CREATE TABLE product_match_audit (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, product_id INTEGER, family_id INTEGER, source_field TEXT, matched_keyword TEXT, source_excerpt TEXT, match_confidence REAL,FOREIGN KEY(company_id) REFERENCES companies(id), FOREIGN KEY(product_id) REFERENCES products(id), FOREIGN KEY(family_id) REFERENCES product_families(id));

CREATE TABLE products (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, family_id INTEGER, parent_product_id INTEGER, hierarchy_level INTEGER, FOREIGN KEY(family_id) REFERENCES product_families(id), FOREIGN KEY(parent_product_id) REFERENCES products(id));

CREATE VIEW v_company_products AS SELECT c.id, c.name, c.exporter_type, c.website, c.state, c.countries_served, c.profile_url, GROUP_CONCAT(DISTINCT pf.name) AS product_families, GROUP_CONCAT(DISTINCT CASE WHEN p.hierarchy_level = 2 THEN p.name END) AS product_variants FROM companies c LEFT JOIN company_products cp ON c.id = cp.company_id LEFT JOIN products p ON cp.product_id = p.id LEFT JOIN product_families pf ON p.family_id = pf.id GROUP BY c.id;

CREATE VIEW v_company_profile AS SELECT c.id, c.name, c.exporter_type, c.website, c.address, c.state, c.postcode, c.countries_served, c.profile_url, GROUP_CONCAT(DISTINCT cert.name) AS certifications, GROUP_CONCAT(DISTINCT acc.name) AS accreditations FROM companies c LEFT JOIN company_certifications cc ON c.id = cc.company_id LEFT JOIN certifications cert ON cc.certification_id = cert.id LEFT JOIN company_accreditations ca ON c.id = ca.company_id LEFT JOIN accreditations acc ON ca.accreditation_id = acc.id GROUP BY c.id;

CREATE VIEW v_product_hierarchy AS SELECT pf.name AS family_name, root.name AS root_product, child.name AS variant_product, child.hierarchy_level FROM product_families pf LEFT JOIN products root ON root.family_id = pf.id AND root.parent_product_id IS NULL LEFT JOIN products child ON child.parent_product_id = root.id ORDER BY pf.sort_order, child.name;