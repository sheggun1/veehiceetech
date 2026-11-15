/* Veehicee scroll reveals: no dependencies; gracefully falls back without IntersectionObserver. */
(() => {
  const root = document.documentElement;
  const reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const header = document.querySelector(".top");
  if (header && "IntersectionObserver" in window) {
    const sentinel = document.createElement("div");
    sentinel.setAttribute("aria-hidden", "true");
    sentinel.style.cssText = "position:absolute;width:1px;height:1px;top:90px;left:0;pointer-events:none";
    document.body.prepend(sentinel);
    const headerObserver = new IntersectionObserver(([entry]) => header.classList.toggle("motion-scrolled", !entry.isIntersecting), {threshold:0});
    headerObserver.observe(sentinel);
  }
  if (reduceMotion || !("IntersectionObserver" in window)) return;
  const groups = [
    [".hero > *", "motion-reveal"],
    [".section .wrap > .section-kicker, .section .wrap > h2, .section .wrap > .intro, .section .wrap > .split, .section .wrap > .chips, .section .wrap > .portfolio-note", "motion-reveal"],
    [".section .cards > *, .section .products > *, .section #products > *, .vicora-inner > *", "motion-reveal"],
    [".section .contact > *", "motion-reveal"]
  ];
  const seen = new Set();
  groups.forEach(([selector, className]) => {
    document.querySelectorAll(selector).forEach((el, index) => {
      if (seen.has(el) || el.closest("#adminModal,.modal,.studio")) return;
      seen.add(el);
      el.classList.add(className);
      if (el.matches(".hero > :first-child")) el.classList.add("motion-from-left");
      if (el.matches(".hero > :last-child")) el.classList.add("motion-from-right");
      const siblings = el.parentElement ? Array.from(el.parentElement.children).filter(node => node.classList.contains(className)) : [];
      el.style.setProperty("--motion-delay", Math.min(siblings.indexOf(el), 5) * 90 + "ms");
      if (el.matches(".card,.prod,.project-card")) el.classList.add("motion-scale");
    });
  });
  root.classList.add("motion-ready");
  const observer = new IntersectionObserver((entries, obs) => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      entry.target.classList.add("is-visible");
      obs.unobserve(entry.target);
    });
  }, {threshold:0.12,rootMargin:"0px 0px -36px 0px"});
  document.querySelectorAll(".motion-reveal").forEach(el => observer.observe(el));
  // Make the opening content visible immediately, avoiding a blank first paint.
  document.querySelectorAll(".hero .motion-reveal").forEach((el, index) => {
    if (index < 2) el.classList.add("is-visible");
  });
})();
