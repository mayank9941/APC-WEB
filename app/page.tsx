"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";

interface PlazaItem {
  code: number;
  name: string;
  has_data: boolean;
}

export default function SearchPage() {
  const router = useRouter();
  const [plazas, setPlazas] = useState<PlazaItem[]>([]);
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const boxRef = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    fetch("/api/plazas")
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json();
      })
      .then(setPlazas)
      .catch((e) => setError(`Could not load plaza list: ${e.message}`));
  }, []);

  useEffect(() => {
    const onClick = (e: MouseEvent) => {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return plazas;
    return plazas.filter(
      (p) => p.name.toLowerCase().includes(q) || String(p.code).includes(q)
    );
  }, [plazas, query]);

  useEffect(() => {
    const el = listRef.current?.children[active] as HTMLElement | undefined;
    el?.scrollIntoView({ block: "nearest" });
  }, [active]);

  const select = (p: PlazaItem) => {
    setOpen(false);
    router.push(`/plaza/${p.code}`);
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (!open || filtered.length === 0) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActive((a) => Math.min(a + 1, filtered.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActive((a) => Math.max(a - 1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      select(filtered[active]);
    } else if (e.key === "Escape") {
      setOpen(false);
    }
  };

  return (
    <main className="search-screen">
      <h1>APC Calculator</h1>
      <p className="subtitle">
        Annual Potential Compensation for point-based toll plazas
      </p>
      <div className="search-box" ref={boxRef}>
        <input
          type="text"
          placeholder="Search plaza by name or code…"
          value={query}
          onChange={(e) => {
            setQuery(e.target.value);
            setOpen(true);
            setActive(0);
          }}
          onFocus={() => {
            setOpen(true);
            setActive(0);
          }}
          onKeyDown={onKeyDown}
          aria-label="Search plaza by name or code"
          autoComplete="off"
        />
        {open && (
          <div className="search-dropdown" ref={listRef}>
            {error && <div className="empty">{error}</div>}
            {!error && filtered.length === 0 && (
              <div className="empty">
                {plazas.length === 0 ? "Loading plazas…" : "No plaza matches your search."}
              </div>
            )}
            {!error &&
              filtered.map((p, i) => (
                <div
                  key={p.code}
                  className={`option${i === active ? " active" : ""}${p.has_data ? "" : " no-data"}`}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    select(p);
                  }}
                  onMouseEnter={() => setActive(i)}
                >
                  <span className="name">
                    {p.name} ({p.code})
                  </span>
                  {!p.has_data && <span className="code">no data</span>}
                </div>
              ))}
          </div>
        )}
      </div>
    </main>
  );
}
