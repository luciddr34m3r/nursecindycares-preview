/* Nurse Cindy — Main JS */

(function () {
  // Scrolled header state
  const header = document.querySelector('.site-header');
  if (header) {
    window.addEventListener('scroll', () => {
      header.classList.toggle('scrolled', window.scrollY > 20);
    });
  }

  // Mobile nav toggle
  const toggle = document.querySelector('.nav-toggle');
  const navLinks = document.querySelector('.nav-links');
  if (toggle && navLinks) {
    toggle.addEventListener('click', () => {
      const open = toggle.classList.toggle('open');
      navLinks.classList.toggle('open', open);
      toggle.setAttribute('aria-expanded', open);
    });

    // Close on link click
    navLinks.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', () => {
        toggle.classList.remove('open');
        navLinks.classList.remove('open');
        toggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Active nav link
  const currentPath = window.location.pathname.replace(/\/$/, '') || '/index.html';
  document.querySelectorAll('.nav-links a').forEach(a => {
    const href = a.getAttribute('href').replace(/\/$/, '');
    if (currentPath.endsWith(href) || (currentPath === '' && href === 'index.html')) {
      a.classList.add('active');
    }
  });

  // Simple contact form → mailto fallback
  const form = document.getElementById('contact-form');
  if (form) {
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const name    = form.querySelector('[name="name"]').value;
      const email   = form.querySelector('[name="email"]').value;
      const org     = form.querySelector('[name="org"]').value;
      const type    = form.querySelector('[name="type"]').value;
      const message = form.querySelector('[name="message"]').value;

      const subject = encodeURIComponent(`Speaking Inquiry from ${name}${org ? ' — ' + org : ''}`);
      const body = encodeURIComponent(
        `Name: ${name}\nEmail: ${email}\nOrganization: ${org}\nInterested in: ${type}\n\nMessage:\n${message}`
      );

      window.location.href = `mailto:cindy.callingallnurses@gmail.com?subject=${subject}&body=${body}`;
    });
  }
})();
