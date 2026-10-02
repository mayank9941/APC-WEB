import Image from "next/image";
import Link from "next/link";
import ihmclLogo from "@/public/ihmcl-logo.png";
import nhaiLogo from "@/public/nhai-logo.png";

export function SiteHeader() {
  return (
    <header className="site-header">
      <div className="site-header-inner">
        <a
          className="brand"
          href="https://ihmcl.co.in"
          target="_blank"
          rel="noopener noreferrer"
          aria-label="Indian Highways Management Company Limited"
        >
          <Image src={ihmclLogo} alt="IHMCL logo" height={44} priority />
          <span className="brand-text">
            <span className="brand-name">IHMCL</span>
            <span className="brand-sub">Indian Highways Management Company Limited</span>
          </span>
        </a>

        <Link href="/" className="site-title">
          <span className="site-title-main">APC 2.0</span>
          <span className="site-title-sub">Annual Potential Collection · Point Based Plazas</span>
        </Link>

        <a
          className="brand brand-right"
          href="https://nhai.gov.in"
          target="_blank"
          rel="noopener noreferrer"
          aria-label="National Highways Authority of India"
        >
          <span className="brand-text">
            <span className="brand-name">NHAI</span>
            <span className="brand-sub">National Highways Authority of India</span>
          </span>
          <Image src={nhaiLogo} alt="NHAI logo" height={44} priority />
        </a>
      </div>
    </header>
  );
}

export function SiteFooter() {
  return (
    <footer className="site-footer">
      © {new Date().getFullYear()} Indian Highways Management Company Limited · National
      Highways Authority of India
    </footer>
  );
}
