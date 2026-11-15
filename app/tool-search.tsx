"use client";

import { useMemo, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { Search, X, ArrowRight } from "lucide-react";

type SearchTool = {
  slug: string;
  name: string;
  description: string;
};

export default function ToolSearch({ tools }: { tools: SearchTool[] }) {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState("");
  const router = useRouter();

  const matches = useMemo(() => {
    const term = query.trim().toLowerCase();
    if (!term) return [];
    return tools.filter(tool =>
      `${tool.name} ${tool.description} ${tool.slug.replaceAll("-", " ")}`
        .toLowerCase()
        .includes(term)
    ).slice(0, 7);
  }, [query, tools]);

  const submitSearch = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (matches.length > 0) router.push(`/tools/${matches[0].slug}`);
  };

  return (
    <div className="toolSearch">
      <button
        type="button"
        className="toolSearchToggle"
        onClick={() => {
          setIsOpen(open => !open);
          if (isOpen) setQuery("");
        }}
        aria-expanded={isOpen}
        aria-controls="homepage-tool-search"
      >
        {isOpen ? <X size={17} /> : <Search size={17} />}
        {isOpen ? "Close search" : "Search tools"}
      </button>

      {isOpen && (
        <div className="toolSearchPanel" id="homepage-tool-search">
          <form className="toolSearchForm" onSubmit={submitSearch} role="search">
            <Search size={18} aria-hidden="true" />
            <input
              autoFocus
              type="search"
              value={query}
              onChange={event => setQuery(event.target.value)}
              placeholder="Search Merge, Rotate, PDF to JPG..."
              aria-label="Search PDF tools"
            />
            <button type="submit" disabled={matches.length === 0}>Go</button>
          </form>
          {query.trim() ? (
            matches.length > 0 ? (
              <div className="toolSearchResults" role="list">
                {matches.map(tool => (
                  <Link
                    href={`/tools/${tool.slug}`}
                    key={tool.slug}
                    className="toolSearchResult"
                    onClick={() => setIsOpen(false)}
                    role="listitem"
                  >
                    <span><b>{tool.name}</b><small>{tool.description}</small></span>
                    <ArrowRight size={16} />
                  </Link>
                ))}
              </div>
            ) : (
              <p className="toolSearchEmpty">No matching tool found. Try another keyword.</p>
            )
          ) : (
            <p className="toolSearchHint">Type a tool name or what you want to do. Select a result to open it directly.</p>
          )}
        </div>
      )}
    </div>
  );
}
