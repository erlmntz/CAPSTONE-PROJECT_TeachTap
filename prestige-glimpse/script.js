// ========== SCROLL REVEAL ANIMATION ==========
const revealElements = document.querySelectorAll('.reveal');

const revealObserver = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('in');
        revealObserver.unobserve(entry.target);
      }
    });
  },
  { threshold: 0.1 }
);

revealElements.forEach((el) => revealObserver.observe(el));

// ========== GENERATE FLOATING PARTICLES ==========
const particlesContainer = document.getElementById('px-particles');
if (particlesContainer) {
  for (let i = 0; i < 24; i++) {
    const particle = document.createElement('div');
    particle.className = 'particle';
    const size = Math.random() * 2 + 1;
    particle.style.cssText = `
      left: ${Math.random() * 100}%;
      top: ${Math.random() * 100}%;
      width: ${size}px;
      height: ${size}px;
      --d: ${(Math.random() * 5 + 3).toFixed(1)}s;
      --dl: ${(Math.random() * 5).toFixed(1)}s;
    `;
    particlesContainer.appendChild(particle);
  }
}

// ========== PARALLAX & MOUSE TILT EFFECTS ==========
const heroSection = document.getElementById('hero-section');
const pxPattern = document.getElementById('px-pattern');
const pxRings = document.getElementById('px-rings');
const pxContent = document.getElementById('px-content');
const banner1 = document.getElementById('px-b1');
const banner2 = document.getElementById('px-b2');

let ticking = false;

function updateParallax() {
  if (ticking) return;
  ticking = true;

  requestAnimationFrame(() => {
    const scrollY = window.scrollY;

    // Hero parallax
    if (heroSection) {
      const heroHeight = heroSection.offsetHeight;
      if (scrollY < heroHeight * 1.2) {
        if (pxPattern) pxPattern.style.transform = `translateY(${scrollY * 0.12}px)`;
        if (pxRings) {
          pxRings.style.transform = `translate(-50%, calc(-50% + ${scrollY * 0.28}px)) rotate(${scrollY * 0.015}deg)`;
        }
        if (pxContent) {
          pxContent.style.transform = `translateY(${scrollY * 0.38}px)`;
          pxContent.style.opacity = Math.max(0, 1 - scrollY / (heroHeight * 0.6));
        }
      }
    }

    // Banner parallax
    [banner1, banner2].forEach((banner) => {
      if (!banner) return;
      const rect = banner.parentElement.getBoundingClientRect();
      const offset = (rect.top + rect.height / 2 - window.innerHeight / 2) * 0.22;
      banner.style.transform = `translateY(${offset}px)`;
    });

    ticking = false;
  });
}

window.addEventListener('scroll', updateParallax, { passive: true });
updateParallax();

// Mouse tilt effect on hero rings and background
if (heroSection) {
  heroSection.addEventListener('mousemove', (e) => {
    const centerX = window.innerWidth / 2;
    const centerY = window.innerHeight / 2;
    const deltaX = (e.clientX - centerX) / centerX;
    const deltaY = (e.clientY - centerY) / centerY;

    if (pxRings) {
      pxRings.style.transform = `translate(calc(-50% + ${deltaX * 10}px), calc(-50% + ${deltaY * 7}px)) rotate(${deltaX * 2.5}deg)`;
    }
    if (pxPattern) {
      pxPattern.style.transform = `translate(${deltaX * 5}px, ${deltaY * 3}px)`;
    }
  });

  heroSection.addEventListener('mouseleave', () => {
    if (pxRings) pxRings.style.transform = 'translate(-50%, -50%)';
    if (pxPattern) pxPattern.style.transform = 'none';
  });
}
