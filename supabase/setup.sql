-- setup.sql — ตั้งค่า Supabase Storage สำหรับเว็บ "หาของในคลัง"
-- วิธีใช้: Supabase > SQL Editor > New query > วางทั้งไฟล์นี้ > Run   (รันซ้ำได้ ไม่พัง)
--
-- แนวคิด:
--   * bucket ชื่อ layout เป็น "สาธารณะ" -> ใครเปิดเว็บก็ดาวน์โหลดไฟล์ผังล่าสุดได้ (ไม่ต้องล็อกอิน)
--   * อัปโหลดได้เฉพาะคนที่ล็อกอิน (ผู้ใช้กลาง 1 คนที่เดย์สร้างไว้) และต้องลงใน versions/ เท่านั้น
--   * ไม่มีสิทธิ์ UPDATE / DELETE เลย -> ไฟล์เก่าไม่ถูกเขียนทับหรือลบ (เก็บประวัติครบ)

-- 1) สร้าง bucket 'layout' (public) จำกัดขนาด 20 MB และเฉพาะไฟล์ .xlsx
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'layout', 'layout', true, 20971520,
  array['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
)
on conflict (id) do update
  set public = true,
      file_size_limit = excluded.file_size_limit,
      allowed_mime_types = excluded.allowed_mime_types;

-- 2) นโยบายความปลอดภัย (RLS) บนตาราง storage.objects
--    ลบของเดิมก่อน เพื่อให้รันซ้ำได้
drop policy if exists "layout_select_all"   on storage.objects;
drop policy if exists "layout_insert_login" on storage.objects;

-- 2.1 อ่าน/ดูรายชื่อไฟล์: ทุกคน (anon) และคนที่ล็อกอิน — เฉพาะ bucket layout
create policy "layout_select_all"
  on storage.objects for select
  to anon, authenticated
  using (bucket_id = 'layout');

-- 2.2 อัปโหลดไฟล์ใหม่: เฉพาะคนที่ล็อกอิน — เฉพาะ bucket layout และโฟลเดอร์ versions/
create policy "layout_insert_login"
  on storage.objects for insert
  to authenticated
  with check (bucket_id = 'layout' and name like 'versions/%');

-- 2.3 ไม่สร้างนโยบาย UPDATE และ DELETE โดยตั้งใจ -> ใครก็แก้/ลบไฟล์เดิมไม่ได้
