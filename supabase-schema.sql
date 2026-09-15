-- ============================================================
-- Schema: รายงานเยี่ยมร้านค้าเทียบแผน (Visit Plan vs Actual Visit)
-- รันไฟล์นี้ใน Supabase: Dashboard > SQL Editor > New query > วางทั้งหมด > Run
-- ============================================================

-- ต้องมี extension นี้เพื่อใช้ gen_random_uuid()
create extension if not exists pgcrypto;

-- ------------------------------------------------------------
-- 1) upload_batches : หนึ่งแถว = การอัปโหลดไฟล์คู่ 1 ครั้ง (1 รอบ/1 เดือน)
-- ------------------------------------------------------------
create table if not exists upload_batches (
  id                uuid primary key default gen_random_uuid(),
  company            text,
  channel            text,
  region             text,
  period_start       date,
  period_end         date,
  plan_filename      text,
  actual_filename    text,
  plan_row_count     int default 0,
  actual_row_count   int default 0,
  uploaded_at        timestamptz not null default now()
);

-- ------------------------------------------------------------
-- 2) visit_plans : ข้อมูลจากไฟล์ "แผนเยี่ยม" (ชีท CallPlanData)
-- ------------------------------------------------------------
create table if not exists visit_plans (
  id                  bigserial primary key,
  batch_id            uuid not null references upload_batches(id) on delete cascade,
  company             text,
  channel             text,
  visit_order         int,
  box_name            text,
  store_code          text not null,
  store_name          text,
  region              text,
  sales_unit          text,
  store_type          text,
  employee_code       text,
  employee_firstname  text,
  employee_lastname   text,
  plan_visit_date     date not null,
  latitude            numeric,
  longitude           numeric,
  created_at          timestamptz not null default now()
);

create index if not exists idx_visit_plans_batch on visit_plans(batch_id);
create index if not exists idx_visit_plans_store_date on visit_plans(store_code, plan_visit_date);

-- ------------------------------------------------------------
-- 3) actual_visits : ข้อมูลจากไฟล์ "รายงานเยี่ยมร้านค้า" (Sheet1)
-- ------------------------------------------------------------
create table if not exists actual_visits (
  id                    bigserial primary key,
  batch_id              uuid not null references upload_batches(id) on delete cascade,
  seq_no                int,
  employee_code         text,
  employee_name         text,
  box_name              text,
  store_code            text not null,
  store_name            text,
  lat_in                numeric,
  lng_in                numeric,
  distance_in_m         numeric,
  time_in               timestamptz,
  reason_in             text,
  lat_out               numeric,
  lng_out               numeric,
  distance_out_m        numeric,
  time_out              timestamptz,
  reason_out            text,
  duration_minutes      int,
  note                  text,
  source_plan_status    text,        -- ค่า "ในแผน/นอกแผน" ดิบจากไฟล์ต้นฉบับ
  visit_date            date not null,
  photo_url             text,
  closed_status         text,
  closed_note           text,
  closed_photo_url      text,
  lat_record            numeric,
  lng_record            numeric,
  distance_record_m     numeric,
  qty_whiskey_white     numeric,
  qty_whiskey_color     numeric,
  qty_rtd               numeric,
  qty_beer              numeric,
  qty_oishi             numeric,
  qty_est               numeric,
  qty_soda              numeric,
  qty_water             numeric,
  qty_other             numeric,
  qty_total             numeric,
  created_at            timestamptz not null default now()
);

create index if not exists idx_actual_visits_batch on actual_visits(batch_id);
create index if not exists idx_actual_visits_store_date on actual_visits(store_code, visit_date);

-- ------------------------------------------------------------
-- Row Level Security
-- เปิด RLS + policy อนุญาต anon key ให้ insert/select ได้
-- (เหมาะกับเครื่องมือใช้ภายในทีม ที่ไม่ต้องแยกสิทธิ์ผู้ใช้)
-- หากต้องการจำกัดสิทธิ์เพิ่มเติม ปรับ policy เหล่านี้ภายหลังได้
-- ------------------------------------------------------------
alter table upload_batches enable row level security;
alter table visit_plans    enable row level security;
alter table actual_visits  enable row level security;

drop policy if exists "anon full access" on upload_batches;
create policy "anon full access" on upload_batches
  for all using (true) with check (true);

drop policy if exists "anon full access" on visit_plans;
create policy "anon full access" on visit_plans
  for all using (true) with check (true);

drop policy if exists "anon full access" on actual_visits;
create policy "anon full access" on actual_visits
  for all using (true) with check (true);
