(function () {
  "use strict";

  var doc = document;
  var root = doc.documentElement;
  var body = doc.body;
  var reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  var finePointer = window.matchMedia("(hover: hover) and (pointer: fine)").matches;
  var supportsIO = "IntersectionObserver" in window;

  root.classList.add("js");

  function $(selector, context) {
    return (context || doc).querySelector(selector);
  }

  function $all(selector, context) {
    return Array.prototype.slice.call((context || doc).querySelectorAll(selector));
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  }

  var header = $(".site-header");
  var nav = $(".nav");
  var navMenu = $("#nav-menu");
  var navToggle = $(".nav__toggle");
  var navIndicator = $(".nav__indicator");
  var navLinks = $all(".nav__menu a[href^='#']");
  var railLinks = $all(".section-rail a[href^='#']");
  var scrollProgress = $(".scroll-progress span");
  var topButton = $(".back-to-top");
  var cursorGlow = $(".cursor-glow");
  var lastScrollY = window.scrollY;
  var ticking = false;
  var activeLink = null;

  var sections = navLinks
    .map(function (link) {
      var target = $(link.getAttribute("href"));
      return target ? { link: link, section: target } : null;
    })
    .filter(Boolean);

  function setMenu(open) {
    if (!navMenu || !navToggle) {
      return;
    }
    navMenu.classList.toggle("is-open", open);
    navToggle.classList.toggle("is-open", open);
    navToggle.setAttribute("aria-expanded", String(open));
  }

  if (navToggle) {
    navToggle.addEventListener("click", function () {
      setMenu(!navMenu.classList.contains("is-open"));
    });
  }

  navLinks.forEach(function (link) {
    link.addEventListener("click", function () {
      setMenu(false);
    });
  });

  function updateNavIndicator(link) {
    if (!nav || !navIndicator || !link) {
      return;
    }
    var navRect = nav.getBoundingClientRect();
    var linkRect = link.getBoundingClientRect();
    nav.style.setProperty("--nav-indicator-left", linkRect.left - navRect.left + "px");
    nav.style.setProperty("--nav-indicator-width", linkRect.width + "px");
    nav.style.setProperty("--nav-indicator-opacity", "1");
  }

  function activateSection(id) {
    navLinks.forEach(function (link) {
      var isActive = link.getAttribute("href") === "#" + id;
      link.classList.toggle("is-active", isActive);
      if (isActive) {
        activeLink = link;
      }
    });
    railLinks.forEach(function (link) {
      link.classList.toggle("is-active", link.getAttribute("href") === "#" + id);
    });
    updateNavIndicator(activeLink);
  }

  if (supportsIO) {
    var navObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            activateSection(entry.target.id);
          }
        });
      },
      { rootMargin: "-36% 0px -56% 0px", threshold: 0.01 }
    );
    sections.forEach(function (item) {
      navObserver.observe(item.section);
    });
  } else if (sections.length) {
    activateSection(sections[0].section.id);
  }

  window.addEventListener("resize", function () {
    updateNavIndicator(activeLink);
  });

  function updateScrollState() {
    var y = window.scrollY;
    var height = Math.max(1, doc.documentElement.scrollHeight - window.innerHeight);
    var percent = clamp((y / height) * 100, 0, 100);
    root.style.setProperty("--scroll-progress", percent.toFixed(2));
    if (scrollProgress) {
      scrollProgress.style.width = percent + "%";
    }
    if (topButton) {
      topButton.classList.toggle("is-visible", y > 520);
      topButton.textContent = y > 520 ? Math.round(percent) + "%" : "↑";
    }
    if (header) {
      header.classList.toggle("is-scrolled", y > 18);
      var scrollingDown = y > lastScrollY && y > 260;
      header.classList.toggle("is-hidden", scrollingDown && !(navMenu && navMenu.classList.contains("is-open")));
    }
    updateTimeline();
    lastScrollY = y;
    ticking = false;
  }

  function requestScrollUpdate() {
    if (!ticking) {
      ticking = true;
      window.requestAnimationFrame(updateScrollState);
    }
  }

  window.addEventListener("scroll", requestScrollUpdate, { passive: true });
  updateScrollState();

  if (topButton) {
    topButton.addEventListener("click", function () {
      window.scrollTo({ top: 0, behavior: reducedMotion ? "auto" : "smooth" });
    });
  }

  if (finePointer && !reducedMotion && cursorGlow) {
    body.classList.add("has-pointer");
    window.addEventListener(
      "pointermove",
      function (event) {
        root.style.setProperty("--cursor-x", event.clientX - 140 + "px");
        root.style.setProperty("--cursor-y", event.clientY - 140 + "px");
      },
      { passive: true }
    );
  }

  doc.addEventListener("visibilitychange", function () {
    body.classList.toggle("is-page-hidden", doc.hidden);
  });

  var revealItems = $all(".reveal, .pipeline, .topology, .code-card");
  if (supportsIO && !reducedMotion) {
    var revealObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) {
            return;
          }
          entry.target.classList.add("is-visible");
          if (entry.target.classList.contains("code-card")) {
            entry.target.classList.add("is-lit");
          }
          revealObserver.unobserve(entry.target);
        });
      },
      { threshold: 0.12 }
    );
    revealItems.forEach(function (item, index) {
      item.style.transitionDelay = Math.min(index % 6, 5) * 45 + "ms";
      revealObserver.observe(item);
    });
  } else {
    revealItems.forEach(function (item) {
      item.classList.add("is-visible");
      item.classList.add("is-lit");
    });
  }

  function animateNumber(element) {
    var target = Number(element.getAttribute("data-target") || "0");
    if (!Number.isFinite(target)) {
      element.textContent = "0";
      return;
    }
    if (reducedMotion) {
      element.textContent = String(target);
      return;
    }
    var duration = 900;
    var startTime = null;
    function tick(time) {
      if (startTime === null) {
        startTime = time;
      }
      var progress = clamp((time - startTime) / duration, 0, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      element.textContent = String(Math.round(target * eased));
      if (progress < 1) {
        window.requestAnimationFrame(tick);
      }
    }
    window.requestAnimationFrame(tick);
  }

  var counters = $all("[data-count]");
  if (supportsIO) {
    var counterObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            animateNumber(entry.target);
            counterObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.42 }
    );
    counters.forEach(function (counter) {
      counterObserver.observe(counter);
    });
  } else {
    counters.forEach(animateNumber);
  }

  if (finePointer && !reducedMotion) {
    $all(".stat-card, .assessment-card, .experiment-card, .issue-card, .code-card, .gallery-card, .summary-card, .info-card, .outcome-card, .pipeline, .node-card").forEach(function (card) {
      card.addEventListener("pointermove", function (event) {
        var rect = card.getBoundingClientRect();
        var x = (event.clientX - rect.left) / rect.width;
        var y = (event.clientY - rect.top) / rect.height;
        var tiltX = (x - 0.5) * 6;
        var tiltY = (0.5 - y) * 6;
        card.style.setProperty("--tilt-x", tiltX.toFixed(2) + "deg");
        card.style.setProperty("--tilt-y", tiltY.toFixed(2) + "deg");
        card.style.setProperty("--mx", x * 100 + "%");
        card.style.setProperty("--my", y * 100 + "%");
      });
      card.addEventListener("pointerleave", function () {
        card.style.setProperty("--tilt-x", "0deg");
        card.style.setProperty("--tilt-y", "0deg");
        card.style.setProperty("--mx", "50%");
        card.style.setProperty("--my", "0%");
      });
    });
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(function (resolve, reject) {
      var textarea = doc.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "");
      textarea.style.position = "fixed";
      textarea.style.top = "-999px";
      body.appendChild(textarea);
      textarea.select();
      try {
        var ok = doc.execCommand("copy");
        body.removeChild(textarea);
        ok ? resolve() : reject(new Error("copy command failed"));
      } catch (error) {
        body.removeChild(textarea);
        reject(error);
      }
    });
  }

  function setCopyState(button, ok) {
    button.textContent = ok ? "已复制" : "请手动复制";
    button.classList.toggle("is-copied", ok);
    button.classList.toggle("is-failed", !ok);
    window.setTimeout(function () {
      button.textContent = button.getAttribute("data-default-label") || "复制";
      button.classList.remove("is-copied", "is-failed");
    }, ok ? 1400 : 1800);
  }

  $all("[data-copy]").forEach(function (button) {
    button.setAttribute("data-default-label", button.textContent.trim() || "复制");
    button.addEventListener("click", function () {
      var card = button.closest(".code-card");
      var code = card ? $("code", card) : null;
      if (!code) {
        return;
      }
      copyText(code.textContent).then(
        function () { setCopyState(button, true); },
        function () { setCopyState(button, false); }
      );
    });
  });

  $all(".code-card").forEach(function (card) {
    var pre = $(".code-block", card);
    var header = $(".code-card__header", card);
    if (!pre || !header) {
      return;
    }
    if (pre.textContent.split("\n").length > 8) {
      card.classList.add("is-foldable", "is-folded");
      var fold = doc.createElement("button");
      fold.type = "button";
      fold.className = "code-fold";
      fold.textContent = "显示全部";
      fold.setAttribute("aria-expanded", "false");
      header.appendChild(fold);
      fold.addEventListener("click", function () {
        var open = card.classList.toggle("is-folded") === false;
        fold.textContent = open ? "收起" : "显示全部";
        fold.setAttribute("aria-expanded", String(open));
      });
    }
  });

  $all(".experiment-card details").forEach(function (details) {
    var summary = $("summary", details);
    if (!summary) {
      return;
    }
    summary.setAttribute("role", "button");
    summary.setAttribute("aria-expanded", String(details.open));
    details.addEventListener("toggle", function () {
      summary.setAttribute("aria-expanded", String(details.open));
    });
  });

  var experimentList = $(".experiment-list");
  var experimentCards = $all(".experiment-card");
  function updateTimeline() {
    if (!experimentList || !experimentCards.length) {
      return;
    }
    var listRect = experimentList.getBoundingClientRect();
    var viewportCenter = window.innerHeight * 0.55;
    var progress = clamp((viewportCenter - listRect.top) / Math.max(1, listRect.height), 0, 1);
    experimentList.style.setProperty("--timeline-progress", (progress * 100).toFixed(2) + "%");
    var current = null;
    experimentCards.forEach(function (card) {
      var rect = card.getBoundingClientRect();
      if (rect.top < viewportCenter && rect.bottom > viewportCenter * 0.55) {
        current = card;
      }
    });
    experimentCards.forEach(function (card) {
      card.classList.toggle("is-current", card === current);
    });
  }

  $all(".issue-card").forEach(function (card) {
    if (!$(".issue-card__status", card)) {
      var status = doc.createElement("span");
      status.className = "issue-card__status";
      status.textContent = "ERROR";
      card.insertBefore(status, card.firstElementChild);
    }
    var dl = $("dl", card);
    if (!dl) {
      return;
    }
    card.classList.add("is-collapsed");
    var toggle = doc.createElement("button");
    toggle.type = "button";
    toggle.className = "issue-toggle";
    toggle.textContent = "查看解决过程";
    toggle.setAttribute("aria-expanded", "false");
    card.appendChild(toggle);
    toggle.addEventListener("click", function () {
      var collapsed = card.classList.toggle("is-collapsed");
      card.classList.toggle("is-resolved", !collapsed);
      toggle.textContent = collapsed ? "查看解决过程" : "收起解决过程";
      toggle.setAttribute("aria-expanded", String(!collapsed));
      var status = $(".issue-card__status", card);
      if (status) {
        status.textContent = collapsed ? "ERROR" : "RESOLVED";
      }
    });
  });

  function outcomeTags(text) {
    if (/192\.168/.test(text)) {
      return ["Ubuntu Server", "Nginx", "SSH", "rsync", "实验网络"];
    }
    return ["HTTPS", "GitHub Actions", "gh-pages", "GitHub Pages"];
  }

  $all(".outcome-card").forEach(function (card) {
    var url = $(".outcome-url", card);
    var action = $(".btn", card);
    if (!url || $(".url-copy-btn", card)) {
      return;
    }
    var copyButton = doc.createElement("button");
    copyButton.type = "button";
    copyButton.className = "url-copy-btn";
    copyButton.textContent = "复制 URL";
    copyButton.setAttribute("data-default-label", "复制 URL");
    copyButton.setAttribute("aria-label", "复制站点地址");
    url.insertAdjacentElement("afterend", copyButton);
    copyButton.addEventListener("click", function () {
      copyText(url.textContent.trim()).then(
        function () { setCopyState(copyButton, true); },
        function () { setCopyState(copyButton, false); }
      );
    });

    var tags = doc.createElement("div");
    tags.className = "outcome-tools";
    tags.innerHTML = outcomeTags(url.textContent).map(function (tag) {
      return "<span>" + tag + "</span>";
    }).join("");
    copyButton.insertAdjacentElement("afterend", tags);

    var preview = doc.createElement("div");
    preview.className = "outcome-preview";
    preview.innerHTML = '<div class="outcome-preview__bar"><span></span><span></span><span></span></div><div class="outcome-preview__body">' + url.textContent.trim() + '</div>';
    if (action) {
      action.insertAdjacentElement("beforebegin", preview);
    } else {
      card.appendChild(preview);
    }
  });

  function galleryCategory(figure) {
    var text = (figure.textContent + " " + (figure.getAttribute("data-experiment") || "") + " " + (figure.getAttribute("data-assessment") || "")).toLowerCase();
    var categories = [];
    if (/典型问题|error|错误|无法|失败|404|disk|timeout/.test(text)) {
      categories.push("error");
    }
    if (/git|github/.test(text)) {
      categories.push("git");
    }
    if (/ssh|scp|rsync|远程/.test(text)) {
      categories.push("ssh");
    }
    if (/nginx|pages|actions|部署|公网|私网|站点/.test(text)) {
      categories.push("deploy");
    }
    if (/linux|ubuntu|virtualbox|node|netplan|bashrc|环境变量|拓扑/.test(text)) {
      categories.push("linux");
    }
    if (/vscode|markdown|hexo|node\.js|npm|开发环境/.test(text)) {
      categories.push("dev");
    }
    if (!categories.length) {
      categories.push("dev");
    }
    return categories.join(" ");
  }

  var galleryCards = $all(".gallery-card");
  galleryCards.forEach(function (figure) {
    figure.setAttribute("data-category", galleryCategory(figure));
    var img = $("img", figure);
    if (img) {
      img.loading = "lazy";
      img.decoding = "async";
    }
  });

  $all(".gallery-filters button").forEach(function (button) {
    button.addEventListener("click", function () {
      var filter = button.getAttribute("data-filter");
      $all(".gallery-filters button").forEach(function (item) {
        item.classList.toggle("is-active", item === button);
      });
      galleryCards.forEach(function (figure) {
        var categories = figure.getAttribute("data-category") || "";
        var visible = filter === "all" || categories.indexOf(filter) !== -1;
        figure.classList.toggle("is-hidden", !visible);
      });
    });
  });

  var lightbox = $("#lightbox");
  var lightboxImage = lightbox ? $(".lightbox__image", lightbox) : null;
  var lightboxCounter = lightbox ? $(".lightbox__counter", lightbox) : null;
  var lightboxTitle = lightbox ? $(".lightbox__title", lightbox) : null;
  var lightboxCaption = lightbox ? $(".lightbox__caption", lightbox) : null;
  var lightboxMeta = lightbox ? $(".lightbox__meta", lightbox) : null;
  var triggers = $all(".lightbox-trigger");
  var activeIndex = 0;
  var lastFocus = null;
  var touchStartX = 0;

  function triggerData(trigger) {
    var figure = trigger.closest("figure");
    var img = $("img", trigger);
    return {
      src: img ? (img.currentSrc || img.src) : "",
      alt: img ? img.alt : "",
      title: figure && $("figcaption strong", figure) ? $("figcaption strong", figure).textContent.trim() : "",
      caption: figure && $("figcaption span", figure) ? $("figcaption span", figure).textContent.trim() : "",
      meta: figure && $("figcaption small", figure) ? $("figcaption small", figure).textContent.trim() : ""
    };
  }

  function showImage(index) {
    if (!lightbox || !lightboxImage || !triggers.length) {
      return;
    }
    activeIndex = (index + triggers.length) % triggers.length;
    var data = triggerData(triggers[activeIndex]);
    lightboxImage.classList.remove("is-switching");
    lightboxImage.src = data.src;
    lightboxImage.alt = data.alt;
    if (lightboxCounter) {
      lightboxCounter.textContent = activeIndex + 1 + " / " + triggers.length;
    }
    if (lightboxTitle) {
      lightboxTitle.textContent = data.title;
    }
    if (lightboxCaption) {
      lightboxCaption.textContent = data.caption;
    }
    if (lightboxMeta) {
      lightboxMeta.textContent = data.meta;
    }
    window.requestAnimationFrame(function () {
      lightboxImage.classList.add("is-switching");
    });
  }

  function openLightbox(index) {
    if (!lightbox) {
      return;
    }
    lastFocus = doc.activeElement;
    lightbox.hidden = false;
    lightbox.classList.add("is-open");
    body.style.overflow = "hidden";
    showImage(index);
    var close = $("[data-close]", lightbox);
    if (close) {
      close.focus();
    }
  }

  function closeLightbox() {
    if (!lightbox || lightbox.hidden) {
      return;
    }
    lightbox.classList.remove("is-open");
    lightbox.hidden = true;
    body.style.overflow = "";
    if (lightboxImage) {
      lightboxImage.src = "";
    }
    if (lastFocus && typeof lastFocus.focus === "function") {
      lastFocus.focus();
    }
  }

  triggers.forEach(function (trigger, index) {
    trigger.addEventListener("click", function () {
      openLightbox(index);
    });
  });

  if (lightbox) {
    $all("[data-close]", lightbox).forEach(function (button) {
      button.addEventListener("click", closeLightbox);
    });
    var prev = $("[data-prev]", lightbox);
    var next = $("[data-next]", lightbox);
    if (prev) {
      prev.addEventListener("click", function () {
        showImage(activeIndex - 1);
      });
    }
    if (next) {
      next.addEventListener("click", function () {
        showImage(activeIndex + 1);
      });
    }
    lightbox.addEventListener("touchstart", function (event) {
      touchStartX = event.changedTouches[0].clientX;
    }, { passive: true });
    lightbox.addEventListener("touchend", function (event) {
      var dx = event.changedTouches[0].clientX - touchStartX;
      if (Math.abs(dx) > 44) {
        showImage(activeIndex + (dx < 0 ? 1 : -1));
      }
    }, { passive: true });
  }

  doc.addEventListener("keydown", function (event) {
    if (!lightbox || lightbox.hidden) {
      return;
    }
    if (event.key === "Escape") {
      closeLightbox();
    } else if (event.key === "ArrowLeft") {
      showImage(activeIndex - 1);
    } else if (event.key === "ArrowRight") {
      showImage(activeIndex + 1);
    }
  });
})();
