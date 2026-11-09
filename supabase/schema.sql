-- Price tracker schema. Run once in the Supabase SQL editor when you switch from CSV.

create table if not exists public.price_checks (
    id            bigint generated always as identity primary key,
    checked_at    timestamptz not null,
    run_id        text        not null,
    product_id    text        not null,
    product_name  text,
    url           text        not null,
    status        text        not null check (status in ('ok', 'failed', 'blocked')),
    price         numeric(12, 2),
    currency      char(3),
    in_stock      boolean,
    method        text,
    used_proxy    boolean     not null default false,
    rendered      boolean     not null default false,
    error         text,
    unique (product_id, checked_at)
);

create index if not exists price_checks_product_time_idx
    on public.price_checks (product_id, checked_at desc);

-- Only the service role key (used by the GitHub Action) can write.
-- Add read policies later if a dashboard reads with the anon key.
alter table public.price_checks enable row level security;

-- Latest successful price per product, with the previous one for change detection.
create or replace view public.latest_prices
with (security_invoker = true) as
select distinct on (product_id)
    product_id,
    product_name,
    url,
    price,
    currency,
    in_stock,
    checked_at,
    lag_price as previous_price,
    case when lag_price is not null and lag_price <> 0
         then round((price - lag_price) / lag_price * 100, 2) end as change_pct
from (
    select *,
           lag(price) over (partition by product_id order by checked_at) as lag_price
    from public.price_checks
    where status = 'ok'
) t
order by product_id, checked_at desc;
