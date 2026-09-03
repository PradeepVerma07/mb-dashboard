-- OPTIONAL migration helper if the old WordPress plugin tables are in the SAME MySQL database
-- Default WordPress table prefix assumed: wp_. Change wp_scooh_ below if your prefix is different.
-- Run schema.sql first. Back up the database before running this file.
INSERT IGNORE INTO sites SELECT * FROM wp_scooh_sites;
INSERT IGNORE INTO clients SELECT * FROM wp_scooh_clients;
INSERT IGNORE INTO campaigns SELECT * FROM wp_scooh_campaigns;
INSERT IGNORE INTO validations SELECT * FROM wp_scooh_validations;
INSERT IGNORE INTO electricity SELECT * FROM wp_scooh_electricity;
INSERT IGNORE INTO vendors SELECT * FROM wp_scooh_vendors;
INSERT IGNORE INTO vendor_jobs SELECT * FROM wp_scooh_vendor_jobs;
INSERT IGNORE INTO proposals SELECT * FROM wp_scooh_proposals;
INSERT IGNORE INTO proposal_sites SELECT * FROM wp_scooh_proposal_sites;
-- invoices has one extra standalone record_status column; map the common plugin fields explicitly.
INSERT IGNORE INTO invoices (id,campaign_id,client_id,requested_date,invoice_no,invoice_date,invoice_amount,invoice_status,hard_copy_required,hard_copy_status,courier_name,tracking_number,dispatch_date,delivered_date,payment_status,payment_date,notes,created_by,updated_by,created_at,updated_at)
SELECT id,campaign_id,client_id,requested_date,invoice_no,invoice_date,invoice_amount,invoice_status,hard_copy_required,hard_copy_status,courier_name,tracking_number,dispatch_date,delivered_date,payment_status,payment_date,notes,created_by,updated_by,created_at,updated_at FROM wp_scooh_invoices;
INSERT IGNORE INTO attachments SELECT * FROM wp_scooh_attachments;
INSERT IGNORE INTO notifications SELECT * FROM wp_scooh_notifications;
INSERT IGNORE INTO activity_log SELECT * FROM wp_scooh_activity_log;
