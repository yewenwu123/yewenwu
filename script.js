const siteConfig = window.siteConfig || {};
const products = window.products || [];

const setText = (id, value) => {
  const node = document.getElementById(id);
  if (node && value) {
    node.textContent = value;
  }
};

const setHtmlTitle = () => {
  if (siteConfig.brandName) {
    document.title = `${siteConfig.brandName} | China Eyewear Sourcing for Australia`;
  }
};

const applySiteConfig = () => {
  setHtmlTitle();
  setText("brandName", siteConfig.brandName);
  setText("brandTagline", siteConfig.brandTagline);
  setText("heroTitle", siteConfig.heroTitle);
  setText("heroText", siteConfig.heroText);
  setText("aboutHeading", siteConfig.aboutHeading);
  setText("aboutIntro", siteConfig.aboutIntro);
  setText("aboutBody", siteConfig.aboutBody);
  setText("contactName", siteConfig.contactName);
  setText("contactRole", siteConfig.contactRole);
  setText("footerBrand", siteConfig.brandName);

  const emailLink = document.getElementById("contactEmail");
  if (emailLink && siteConfig.contactEmail) {
    emailLink.href = `mailto:${siteConfig.contactEmail}`;
    emailLink.textContent = siteConfig.contactEmail;
  }

  const whatsappLink = document.getElementById("contactWhatsapp");
  if (whatsappLink && siteConfig.whatsappLink && siteConfig.whatsappLabel) {
    whatsappLink.href = siteConfig.whatsappLink;
    whatsappLink.textContent = siteConfig.whatsappLabel;
  }

  const form = document.getElementById("consultForm");
  const formStatus = document.getElementById("formStatus");
  if (!form || !formStatus) {
    return;
  }

  if (siteConfig.formEndpoint) {
    form.action = siteConfig.formEndpoint;
    formStatus.textContent = siteConfig.formStatusReady || "Form endpoint connected. Test it once before publishing ads.";
  } else {
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      formStatus.textContent = "Add your form endpoint in data/site-config.js before publishing this page.";
    });
  }
};

const renderProducts = () => {
  const grid = document.getElementById("productGrid");
  if (!grid) {
    return;
  }

  grid.innerHTML = products.map((product) => `
    <article class="product-card">
      <div class="product-image">
        <img src="${product.image}" alt="${product.title}">
      </div>
      <div class="product-meta">
        <span class="pill">${product.category}</span>
        <span class="pill">${product.audience}</span>
      </div>
      <div>
        <h3>${product.title}</h3>
        <p>${product.description}</p>
      </div>
      <a class="button button-secondary enquire-button" href="#contact" data-style="${product.title}">Ask about this style</a>
    </article>
  `).join("");

  grid.querySelectorAll(".enquire-button").forEach((button) => {
    button.addEventListener("click", () => {
      const field = document.getElementById("styleInterest");
      if (field) {
        field.value = button.dataset.style || "";
      }
    });
  });
};

document.getElementById("currentYear").textContent = new Date().getFullYear();
applySiteConfig();
renderProducts();
