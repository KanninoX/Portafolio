// Año en el footer
document.getElementById('year').textContent = new Date().getFullYear();

// Menú móvil
const navToggle = document.getElementById('navToggle');
const navLinks = document.getElementById('navLinks');
navToggle.addEventListener('click', () => {
  const isOpen = navLinks.classList.toggle('is-open');
  navToggle.setAttribute('aria-expanded', isOpen);
});
navLinks.querySelectorAll('a').forEach(link => {
  link.addEventListener('click', () => {
    navLinks.classList.remove('is-open');
    navToggle.setAttribute('aria-expanded', 'false');
  });
});

// Revelado al hacer scroll
const revealEls = document.querySelectorAll('.reveal');
const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

if (prefersReducedMotion) {
  revealEls.forEach(el => el.classList.add('is-visible'));
} else if ('IntersectionObserver' in window) {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('is-visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
  revealEls.forEach(el => observer.observe(el));
} else {
  revealEls.forEach(el => el.classList.add('is-visible'));
}

// Efecto de máquina de escribir en la terminal del hero
const terminalLines = [
  { prompt: '$ whoami', delay: 30 },
  { prompt: '> Ramón Osorio — Encargado de Remuneraciones @ Invermar S.A.', delay: 12 },
  { prompt: '$ cat trayectoria.log', delay: 30 },
  { prompt: '> 15+ años en nómina y RR.HH. · industria salmonera y construcción', delay: 12 },
  { prompt: '$ ./aprender-a-programar.sh --stack=AppsScript,Python,SQL,Java', delay: 18 },
  { prompt: '> Automatizando lo que antes se hacía a mano, una línea a la vez.', delay: 12 },
];

const terminalBody = document.getElementById('terminalBody');

function typeTerminal() {
  if (prefersReducedMotion) {
    terminalBody.textContent = terminalLines.map(l => l.prompt).join('\n') + '\n$ _';
    return;
  }

  let lineIndex = 0;
  let charIndex = 0;
  let output = '';

  function typeChar() {
    if (lineIndex >= terminalLines.length) {
      terminalBody.textContent = output + '\n$ ';
      blinkCursor();
      return;
    }
    const line = terminalLines[lineIndex];
    if (charIndex === 0 && lineIndex > 0) output += '\n';

    if (charIndex < line.prompt.length) {
      output += line.prompt[charIndex];
      terminalBody.textContent = output + '▌';
      charIndex++;
      setTimeout(typeChar, line.delay);
    } else {
      charIndex = 0;
      lineIndex++;
      setTimeout(typeChar, 220);
    }
  }
  typeChar();
}

function blinkCursor() {
  let visible = true;
  setInterval(() => {
    visible = !visible;
    terminalBody.textContent = terminalBody.textContent.replace(/[▌ ]$/, '') + (visible ? '▌' : ' ');
  }, 600);
}

// Inicia la animación cuando la terminal entra en pantalla
if ('IntersectionObserver' in window) {
  const termObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        typeTerminal();
        termObserver.disconnect();
      }
    });
  }, { threshold: 0.3 });
  termObserver.observe(document.querySelector('.hero-terminal'));
} else {
  typeTerminal();
}
