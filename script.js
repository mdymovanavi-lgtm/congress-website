(function () {
  'use strict';

  /* ---------- Предзагрузка внутренних страниц (только десктоп с мышью) ---------- */
  if (window.matchMedia('(hover: hover) and (pointer: fine)').matches) {
    var prefetched = Object.create(null);
    document.addEventListener('mouseover', function (e) {
      var a = e.target.closest('a[href]');
      if (!a) return;
      var href = a.getAttribute('href');
      if (!href || href.charAt(0) === '#') return;
      if (/^(mailto:|tel:|javascript:)/i.test(href)) return;
      if (/\.(pdf|jpe?g|png|webp|zip|mp4)(\?|#|$)/i.test(href)) return;
      var url;
      try { url = new URL(href, location.href); } catch (err) { return; }
      if (url.origin !== location.origin) return;
      if (url.pathname === location.pathname && url.hash === location.hash) return;
      if (prefetched[url.href]) return;
      prefetched[url.href] = 1;
      var link = document.createElement('link');
      link.rel = 'prefetch';
      link.href = url.href;
      document.head.appendChild(link);
    }, true);
  }

  /* ---------- Мобильное меню ---------- */
  var burger = document.querySelector('.burger');
  var mobileMenu = document.querySelector('.mobile-menu');
  if (burger && mobileMenu) {
    burger.addEventListener('pointerup', function (e) {
      e.preventDefault();
      var open = mobileMenu.hasAttribute('hidden');
      if (open) { mobileMenu.removeAttribute('hidden'); } else { mobileMenu.setAttribute('hidden', ''); }
      burger.classList.toggle('is-open', open);
      burger.setAttribute('aria-expanded', String(open));
    });
    mobileMenu.addEventListener('click', function (e) {
      if (e.target.closest('a')) {
        mobileMenu.setAttribute('hidden', '');
        burger.classList.remove('is-open');
        burger.setAttribute('aria-expanded', 'false');
      }
    });
  }

  /* ---------- Вкладки «Как принять участие» ---------- */
  var applyTabs = document.querySelector('[data-apply-tabs]');
  if (applyTabs) {
    var tabButtons = applyTabs.querySelectorAll('.apply-tab');
    var panels = document.querySelectorAll('.apply-panel');
    applyTabs.addEventListener('click', function (e) {
      var btn = e.target.closest('.apply-tab');
      if (!btn) return;
      var target = btn.dataset.tab;
      tabButtons.forEach(function (b) {
        b.classList.toggle('is-active', b === btn);
        b.setAttribute('aria-selected', String(b === btn));
      });
      panels.forEach(function (p) { p.hidden = p.dataset.panel !== target; });
    });
  }

  /* ---------- Модал «Связаться» ---------- */
  var modal = document.getElementById('contactModal');
  var lastFocus = null;
  var modalFocusables = [];

  function getModalFocusables() {
    if (!modal) return [];
    return Array.prototype.slice.call(
      modal.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), textarea:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])')
    ).filter(function (el) { return el.offsetParent !== null || el === modal.querySelector('.modal__close'); });
  }

  function openModal() {
    if (!modal) return;
    lastFocus = document.activeElement;
    modal.removeAttribute('hidden');
    document.body.style.overflow = 'hidden';
    modalFocusables = getModalFocusables();
    var first = modalFocusables[0];
    if (first) first.focus();
  }
  function closeModal() {
    if (!modal) return;
    modal.setAttribute('hidden', '');
    document.body.style.overflow = '';
    if (lastFocus) lastFocus.focus();
  }
  function handleContactAction(e) {
    if (e.target.closest('.js-open-contact')) {
      e.preventDefault();
      openModal();
    }
    if (e.target.closest('.js-close-contact')) {
      e.preventDefault();
      closeModal();
    }
  }
  document.addEventListener('pointerup', handleContactAction);
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && modal && !modal.hasAttribute('hidden')) closeModal();
    if (e.key === 'Tab' && modal && !modal.hasAttribute('hidden')) {
      modalFocusables = getModalFocusables();
      if (modalFocusables.length < 2) return;
      var first = modalFocusables[0];
      var last = modalFocusables[modalFocusables.length - 1];
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }
  });

  /* ---------- Форма ---------- */
  var form = document.getElementById('contactForm');
  if (form) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!form.reportValidity()) return;
      var success = form.querySelector('.form-success');
      var submit = form.querySelector('.modal__submit');
      var errorEl = form.querySelector('.form-error');
      if (submit) { submit.disabled = true; submit.textContent = 'Отправляем…'; }
      if (errorEl) errorEl.hidden = true;

      var body = new FormData(form);
      if (!body.get('form-name')) body.append('form-name', 'contact');

      fetch('/', {
        method: 'POST',
        headers: { 'Accept': 'application/json' },
        body: body
      }).then(function (res) {
        if (!res.ok) throw new Error('submit failed');
        if (success) success.removeAttribute('hidden');
        if (submit) submit.style.display = 'none';
        form.querySelectorAll('.field, .form-consent, .visually-hidden').forEach(function (el) { el.style.display = 'none'; });
      }).catch(function () {
        if (submit) { submit.disabled = false; submit.textContent = 'Отправить'; }
        if (!errorEl) {
          errorEl = document.createElement('p');
          errorEl.className = 'form-error';
          errorEl.textContent = 'Не удалось отправить. Попробуйте позже или напишите на office@congress.ru';
          form.appendChild(errorEl);
        }
        errorEl.hidden = false;
      });
    });
  }

  /* ---------- Переключатель услуг (главная) ---------- */
  var items = document.querySelectorAll('.services__item');
  if (items.length && items[0].tagName !== 'A') {
    var panelIndex = document.querySelector('[data-panel-index]');
    var panelTitle = document.querySelector('[data-panel-title]');
    var panelDesc = document.querySelector('[data-panel-desc]');
    var panelLink = document.querySelector('[data-panel-link]');
    items.forEach(function (btn, i) {
      btn.addEventListener('click', function () {
        items.forEach(function (b) { b.classList.remove('is-active'); });
        btn.classList.add('is-active');
        if (panelIndex) panelIndex.textContent = String(i + 1).padStart(2, '0');
        if (panelTitle) panelTitle.textContent = btn.textContent;
        if (panelDesc) panelDesc.textContent = btn.dataset.desc || '';
        if (panelLink) panelLink.setAttribute('href', btn.dataset.link || 'services.html');
      });
    });
  }

  /* ---------- Подсветка активного раздела (страница услуг) ---------- */
  var svcLinks = document.querySelectorAll('.svc-nav a');
  var svcSections = document.querySelectorAll('.svc');
  if (svcLinks.length && svcSections.length && 'IntersectionObserver' in window) {
    var byId = {};
    svcLinks.forEach(function (a) { byId[a.getAttribute('href').slice(1)] = a; });
    var spy = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          svcLinks.forEach(function (a) { a.classList.remove('is-active'); });
          var link = byId[entry.target.id];
          if (link) link.classList.add('is-active');
        }
      });
    }, { rootMargin: '-30% 0px -60% 0px' });
    svcSections.forEach(function (s) { spy.observe(s); });
    svcLinks.forEach(function (a) {
      a.addEventListener('click', function (e) {
        var id = a.getAttribute('href');
        if (!id || id.charAt(0) !== '#') return;
        var target = document.getElementById(id.slice(1));
        if (!target) return;
        e.preventDefault();
        target.scrollIntoView({ behavior: 'auto', block: 'start' });
        if (history.replaceState) history.replaceState(null, '', id);
      });
    });
  }

  /* ---------- Карта: logomark прыгает по городам ---------- */
  var geoPin = document.querySelector('.geo__pin');
  var geoDots = document.querySelectorAll('.geo__dot--target');
  if (geoPin && geoDots.length) {
    var geoMarker = geoPin.querySelector('.geo__pin-marker');
    var cityListItems = document.querySelectorAll('.geo__city-col li[data-city]');
    var geoSection = geoPin.closest('.geo');
    var geoIndex = 0;
    var geoTravelMs = 780;
    var geoBounceMs = 360;
    var geoActive = true;
    var geoRunId = 0;

    function geoIsLive() {
      return geoActive && !document.hidden;
    }

    function stopGeo() {
      geoActive = false;
      geoRunId++;
    }

    function startGeo() {
      if (geoActive) return;
      geoActive = true;
      geoRunId++;
      var runId = geoRunId;
      playAtCity(geoIndex, function (peakLift) {
        if (runId === geoRunId) goToNextCity(peakLift);
      });
    }

    if (geoSection && 'IntersectionObserver' in window) {
      new IntersectionObserver(function (entries) {
        if (entries[0].isIntersecting && !document.hidden) startGeo();
        else stopGeo();
      }, { threshold: 0.08 }).observe(geoSection);
    }
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) stopGeo();
      else if (geoSection) {
        var rect = geoSection.getBoundingClientRect();
        if (rect.bottom > 0 && rect.top < window.innerHeight) startGeo();
      }
    });

    function parseDotPos(dot) {
      return {
        left: parseFloat(dot.style.left),
        top: parseFloat(dot.style.top)
      };
    }

    function setPinPos(left, top) {
      geoPin.style.left = left + '%';
      geoPin.style.top = top + '%';
    }

    function highlightCity(dot) {
      var city = dot.dataset.city;
      geoDots.forEach(function (d) { d.classList.remove('is-lit'); });
      dot.classList.add('is-lit');
      cityListItems.forEach(function (li) {
        li.classList.toggle('is-current', li.dataset.city === city);
      });
    }

    function setMarkerTransform(yOffset, scaleX, scaleY) {
      if (!geoMarker) return;
      geoMarker.style.transform = 'translate(-50%, -100%) translateY(' + yOffset + 'px) scale(' + scaleX + ', ' + scaleY + ')';
    }

    function markerHeight() {
      return geoMarker ? geoMarker.offsetHeight || 28 : 28;
    }

    function animateBounce(heightRatio, stopAtPeak, done) {
      var start = null;
      var h = markerHeight();
      var duration = stopAtPeak ? geoBounceMs * 0.52 : geoBounceMs;
      var runId = geoRunId;

      function frame(ts) {
        if (!geoIsLive() || runId !== geoRunId) return;
        if (!start) start = ts;
        var progress = Math.min((ts - start) / duration, 1);
        var phase = stopAtPeak ? progress * 0.5 : progress;
        var y = -h * heightRatio * Math.sin(phase * Math.PI);
        var sx = 1;
        var sy = 1;

        if (!stopAtPeak && progress > 0.82) {
          var landT = (progress - 0.82) / 0.18;
          var squash = Math.sin(landT * Math.PI);
          sx = 1 + 0.08 * squash;
          sy = 1 - 0.06 * squash;
        } else if (!stopAtPeak) {
          sx = 0.99;
          sy = 1.02;
        }

        setMarkerTransform(y, sx, sy);

        if (progress < 1) {
          requestAnimationFrame(frame);
        } else {
          done(stopAtPeak ? y : 0);
        }
      }

      requestAnimationFrame(frame);
    }

    function animateTravel(fromIdx, toIdx, startLift, done) {
      var from = parseDotPos(geoDots[fromIdx]);
      var to = parseDotPos(geoDots[toIdx]);
      var arcHeight = 8;
      var start = null;
      var runId = geoRunId;

      function frame(ts) {
        if (!geoIsLive() || runId !== geoRunId) return;
        if (!start) start = ts;
        var t = Math.min((ts - start) / geoTravelMs, 1);
        var ease = t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2;
        var x = from.left + (to.left - from.left) * ease;
        var y = from.top + (to.top - from.top) * ease;
        var mapArc = -arcHeight * Math.sin(t * Math.PI);
        var lift = startLift * (1 - t);

        setPinPos(x, y + mapArc);

        if (t > 0.9) {
          var landT = (t - 0.9) / 0.1;
          var squash = Math.sin(landT * Math.PI);
          setMarkerTransform(lift, 1 + 0.1 * squash, 1 - 0.07 * squash);
        } else {
          setMarkerTransform(lift, 0.98, 1.02);
        }

        if (t < 1) {
          requestAnimationFrame(frame);
        } else {
          setPinPos(to.left, to.top);
          setMarkerTransform(0, 1, 1);
          done();
        }
      }

      requestAnimationFrame(frame);
    }

    function playAtCity(cityIdx, done) {
      highlightCity(geoDots[cityIdx]);
      animateBounce(0.5, false, function () {
        animateBounce(0.26, true, function (peakLift) {
          done(peakLift);
        });
      });
    }

    function goToNextCity(peakLift) {
      var fromIdx = geoIndex;
      geoIndex = (geoIndex + 1) % geoDots.length;
      animateTravel(fromIdx, geoIndex, peakLift, function () {
        playAtCity(geoIndex, function (nextPeakLift) {
          goToNextCity(nextPeakLift);
        });
      });
    }

    var startPos = parseDotPos(geoDots[0]);
    setPinPos(startPos.left, startPos.top);

    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      stopGeo();
      highlightCity(geoDots[0]);
    } else {
      var geoVisibleNow = !geoSection || (function () {
        var rect = geoSection.getBoundingClientRect();
        return rect.top < window.innerHeight && rect.bottom > 0;
      }());
      if (geoVisibleNow) {
        playAtCity(0, goToNextCity);
      } else {
        stopGeo();
      }
    }
  }

  /* ---------- Календарь событий: ближайшее крупно + аккордеон остальных ---------- */
  var CAL_EVENTS = [
    { title: 'ПМЭФ-2026', start: '2026-06-03', end: '2026-06-06', city: 'Санкт-Петербург', dates: '3–6 июня 2026', status: 'open', link: 'https://congress.ru', img: 'assets/photo/projects/pmef.webp', desc: 'Подать заявку на участие в проекте в качестве волонтёра можно до 24 марта 2026 года.' },
    { title: 'ПМЮФ-2026', start: '2026-06-24', end: '2026-06-26', city: 'Санкт-Петербург', dates: '24–26 июня 2026', status: 'open', link: 'https://congress.ru', img: null, desc: 'Подать заявку на участие в проекте в качестве волонтёра можно до 3 мая 2026 года.' },
    { title: 'ВЭФ-2026', start: '2026-09-01', end: '2026-09-04', city: 'Владивосток', dates: '1–4 сентября 2026', status: 'open', link: 'calendar/vef-2026.html', img: 'assets/photo/projects/vef.webp', desc: 'Подать заявку на участие в проекте в качестве волонтёра можно до 15 мая 2026 года.' },
    { title: 'Форум объединённых культур', start: '2026-09-24', end: '2026-09-26', city: 'Санкт-Петербург', dates: '24–26 сентября 2026', status: 'open', link: 'calendar/fok-2026.html', img: null, desc: 'Подать заявку на участие в проекте в качестве волонтёра можно до 23 июля 2026 года.' },
    { title: 'РЭН-2026', start: '2026-10-14', end: '2026-10-16', city: 'Москва', dates: '14–16 октября 2026', status: 'open', link: 'calendar/ren-2026.html', img: null, desc: 'Подать заявку на участие в проекте в качестве волонтёра можно ближе к началу набора.' },
    { title: 'КМУ-2026', start: '2026-11-25', end: '2026-11-27', city: 'Сириус', dates: '25–27 ноября 2026', status: 'open', link: 'calendar/kmu-2026.html', img: null, desc: 'Подать заявку на участие в проекте в качестве волонтёра можно ближе к началу набора.' }
  ];

  var calRoot = document.querySelector('[data-calendar]');
  if (calRoot) {
    var featureEl = calRoot.querySelector('[data-cal-feature]');
    var listEl = calRoot.querySelector('[data-cal-list]');
    var restLabel = document.querySelector('[data-cal-rest-label]');
    var calEmpty = document.querySelector('[data-cal-empty]');
    var monthAbbr = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    function statusMarkup(status) {
      var cls = status === 'open' ? 'cal-status--open' : 'cal-status--closed';
      var text = status === 'open' ? 'Запись открыта' : 'Скоро открытие';
      return '<span class="cal-status ' + cls + '">' + text + '</span>';
    }

    function mediaMarkup(ev) {
      if (ev.img) {
        return '<div class="cal-card-full__media"><img src="' + ev.img + '" alt="' + ev.title + '" loading="lazy"></div>';
      }
      return '<div class="cal-card-full__media cal-card-full__media--placeholder" aria-hidden="true"></div>';
    }

    function fullCardMarkup(ev, opts) {
      opts = opts || {};
      var day = String(ev.startDate.getDate()).padStart(2, '0');
      var month = monthAbbr[ev.startDate.getMonth()];
      var eyebrow = opts.eyebrow ? '<p class="cal-card-full__eyebrow">' + opts.eyebrow + '</p>' : '';
      var showTitle = opts.showTitle !== false;
      var showCity = opts.hideCity !== true;
      var showStatus = opts.hideStatus !== true;
      var title = showTitle ? '<p class="cal-card-full__title">' + ev.title + '</p>' : '';
      var status = showStatus ? statusMarkup(ev.status) : '';
      var head = (title || status) ? '<div class="cal-card-full__head">' + title + status + '</div>' : '';
      var city = showCity ? '<span class="cal-card-full__city">' + ev.city + '</span>' : '';
      return '<div class="cal-card-full">' +
        '<div class="cal-card-full__date"><span class="cal-card-full__date-day">' + day + '</span><span class="cal-card-full__date-month">' + month + '</span></div>' +
        '<div class="cal-card-full__info">' + eyebrow + head + city +
          '<p class="cal-card-full__meta">' + ev.dates + '</p>' +
          '<p class="cal-card-full__desc">' + ev.desc + '</p>' +
          '<a class="cal-link" href="' + ev.link + '">Подробнее</a>' +
        '</div>' +
        mediaMarkup(ev) +
      '</div>';
    }

    var upcoming = CAL_EVENTS
      .map(function (ev) {
        return {
          title: ev.title, city: ev.city, dates: ev.dates, status: ev.status, link: ev.link, img: ev.img, desc: ev.desc,
          startDate: new Date(ev.start), endDate: new Date(ev.end)
        };
      })
      .filter(function (ev) { return ev.endDate >= today; })
      .sort(function (a, b) { return a.startDate - b.startDate; });

    if (!upcoming.length) {
      calRoot.hidden = true;
      if (calEmpty) calEmpty.hidden = false;
    } else {
      var featured = upcoming[0];
      var rest = upcoming.slice(1);
      var isOngoing = featured.startDate <= today;
      var eyebrowText = isOngoing ? 'Проходит прямо сейчас' : 'Ближайшее мероприятие';

      featureEl.innerHTML = fullCardMarkup(featured, { eyebrow: eyebrowText });

      if (rest.length) {
        if (restLabel) restLabel.hidden = false;
        rest.forEach(function (ev) {
          var day = String(ev.startDate.getDate()).padStart(2, '0');
          var month = monthAbbr[ev.startDate.getMonth()];
          var item = document.createElement('div');
          item.className = 'cal-item';
          item.innerHTML =
            '<button class="cal-item__toggle" type="button" aria-expanded="false">' +
              '<span class="cal-item__date"><span class="cal-item__date-day">' + day + '</span><span class="cal-item__date-month">' + month + '</span></span>' +
              '<span class="cal-item__body">' +
                '<span class="cal-item__head">' +
                  '<span class="cal-item__title">' + ev.title + '</span>' +
                  statusMarkup(ev.status) +
                '</span>' +
                '<span class="cal-item__city">' + ev.city + '</span>' +
              '</span>' +
              '<span class="cal-item__chevron" aria-hidden="true"></span>' +
            '</button>' +
            '<div class="cal-item__panel" hidden>' + fullCardMarkup(ev, { showTitle: false, hideCity: true }) + '</div>';
          listEl.appendChild(item);
        });

        listEl.addEventListener('click', function (e) {
          var btn = e.target.closest('.cal-item__toggle');
          if (!btn) return;
          var item = btn.closest('.cal-item');
          var wasOpen = item.classList.contains('is-open');
          listEl.querySelectorAll('.cal-item.is-open').forEach(function (openItem) {
            openItem.classList.remove('is-open');
            openItem.querySelector('.cal-item__toggle').setAttribute('aria-expanded', 'false');
            openItem.querySelector('.cal-item__panel').hidden = true;
          });
          if (!wasOpen) {
            item.classList.add('is-open');
            btn.setAttribute('aria-expanded', 'true');
            item.querySelector('.cal-item__panel').hidden = false;
          }
        });
      }
    }
  }

  /* ---------- Появление при скролле ---------- */
  var reveals = document.querySelectorAll('.reveal');
  var reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* Фоновое видео: постер сразу, ролик после отрисовки; на медленной сети — только постер */
  var heroVideo = document.querySelector('.hero__video');
  function shouldLoadHeroVideo() {
    if (!heroVideo || reduced) return false;
    if (window.matchMedia('(max-width: 899px)').matches) return false;
    if (window.matchMedia('(hover: none), (pointer: coarse)').matches) return false;
    var conn = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    if (conn) {
      if (conn.saveData) return false;
      if (conn.effectiveType === 'slow-2g' || conn.effectiveType === '2g' || conn.effectiveType === '3g') return false;
      if (conn.effectiveType === '4g' && typeof conn.downlink === 'number' && conn.downlink < 2.5) return false;
    }
    return window.matchMedia('(hover: hover) and (pointer: fine)').matches;
  }
  function loadHeroVideo() {
    if (!shouldLoadHeroVideo() || heroVideo.dataset.loaded) return;
    var src = heroVideo.getAttribute('data-src');
    if (!src) return;
    heroVideo.dataset.loaded = '1';
    var source = document.createElement('source');
    source.src = src;
    source.type = 'video/mp4';
    heroVideo.appendChild(source);
    heroVideo.playbackRate = 0.5;
    heroVideo.addEventListener('play', function () { heroVideo.playbackRate = 0.5; });
    heroVideo.addEventListener('loadeddata', function () {
      heroVideo.playbackRate = 0.5;
      var playPromise = heroVideo.play();
      if (playPromise && playPromise.catch) playPromise.catch(function () {});
    });
    heroVideo.load();
  }
  if (heroVideo && shouldLoadHeroVideo()) {
    var startHeroVideo = function () {
      if ('requestIdleCallback' in window) {
        requestIdleCallback(loadHeroVideo, { timeout: 1800 });
      } else {
        setTimeout(loadHeroVideo, 250);
      }
    };
    requestAnimationFrame(function () {
      requestAnimationFrame(startHeroVideo);
    });
  }

  /* ---------- Счётчики в блоке цифр ---------- */
  var counters = document.querySelectorAll('.js-count');
  function runCounter(el) {
    var target = parseInt(el.dataset.count, 10) || 0;
    var duration = 1400;
    var start = null;
    function frame(ts) {
      if (!start) start = ts;
      var p = Math.min((ts - start) / duration, 1);
      var eased = 1 - Math.pow(1 - p, 3); /* ease-out cubic */
      el.textContent = Math.round(target * eased);
      if (p < 1) requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);
  }
  if (counters.length) {
    if (reduced) {
      counters.forEach(function (el) { el.textContent = el.dataset.count; });
    } else {
      var checkCounters = function () {
        counters.forEach(function (el) {
          if (el.dataset.done) return;
          var r = el.getBoundingClientRect();
          if (r.top < window.innerHeight - 40 && r.bottom > 0) {
            el.dataset.done = '1';
            runCounter(el);
          }
        });
      };
      window.addEventListener('scroll', checkCounters, { passive: true });
      window.addEventListener('resize', checkCounters, { passive: true });
      checkCounters();
    }
  }
  if (reveals.length && 'IntersectionObserver' in window && !reduced) {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add('is-visible');
          io.unobserve(entry.target);
        }
      });
    }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
    reveals.forEach(function (el) { io.observe(el); });
  } else {
    reveals.forEach(function (el) { el.classList.add('is-visible'); });
  }

  /* ---------- Карусель благодарностей ---------- */
  var thanksWrap = document.querySelector('.trust__thanks-wrap');
  var thanksCarousel = document.querySelector('[data-thanks-carousel]');
  if (thanksCarousel && thanksWrap) {
    var prevBtn = thanksWrap.querySelector('.trust__thanks-arrow--prev');
    var nextBtn = thanksWrap.querySelector('.trust__thanks-arrow--next');
    if (prevBtn && nextBtn) {
      var scrollAmount = 240; // ширина элемента + gap
      prevBtn.addEventListener('click', function () {
        thanksCarousel.scrollBy({ left: -scrollAmount, behavior: 'smooth' });
      });
      nextBtn.addEventListener('click', function () {
        thanksCarousel.scrollBy({ left: scrollAmount, behavior: 'smooth' });
      });
    }
  }

  /* ---------- Горизонтальная карусель «Что делали» (страница проекта) ---------- */
  document.querySelectorAll('[data-proj-stories]').forEach(function (root) {
    var viewport = root.querySelector('.proj-stories__viewport');
    var track = root.querySelector('.proj-stories__track');
    var stories = root.querySelectorAll('.proj-story');
    var section = root.closest('.proj-scope');
    var prev = section ? section.querySelector('.proj-stories__arrow--prev') : null;
    var next = section ? section.querySelector('.proj-stories__arrow--next') : null;
    if (!viewport || !stories.length) return;

    function step() {
      if (!stories[0]) return 0;
      var gap = track ? parseFloat(getComputedStyle(track).gap) || 16 : 16;
      return stories[0].offsetWidth + gap;
    }

    if (stories.length < 2) {
      if (prev) prev.hidden = true;
      if (next) next.hidden = true;
      return;
    }

    if (prev) {
      prev.addEventListener('click', function () {
        viewport.scrollBy({ left: -step(), behavior: 'smooth' });
      });
    }
    if (next) {
      next.addEventListener('click', function () {
        viewport.scrollBy({ left: step(), behavior: 'smooth' });
      });
    }
  });

  /* ---------- Карусель фото проекта ---------- */
  document.querySelectorAll('[data-carousel]').forEach(function (root) {
    var track = root.querySelector('.carousel__track');
    var slides = root.querySelectorAll('.carousel__slide');
    var prev = root.querySelector('.carousel__arrow--prev');
    var next = root.querySelector('.carousel__arrow--next');
    var dotsWrap = root.querySelector('.carousel__dots');
    if (!track || !slides.length) return;
    var index = 0;
    if (slides.length < 2) {
      if (prev) prev.hidden = true;
      if (next) next.hidden = true;
      if (dotsWrap) dotsWrap.hidden = true;
      return;
    }
    var dots = [];
    slides.forEach(function (_, i) {
      var dot = document.createElement('button');
      dot.type = 'button';
      dot.className = 'carousel__dot';
      dot.setAttribute('aria-label', 'Фото ' + (i + 1));
      dot.addEventListener('click', function () { goTo(i); });
      if (dotsWrap) dotsWrap.appendChild(dot);
      dots.push(dot);
    });
    function goTo(i) {
      index = (i + slides.length) % slides.length;
      track.style.transform = 'translateX(-' + (index * 100) + '%)';
      dots.forEach(function (d, di) { d.classList.toggle('is-active', di === index); });
    }
    if (prev) prev.addEventListener('click', function () { goTo(index - 1); });
    if (next) next.addEventListener('click', function () { goTo(index + 1); });
    goTo(0);
  });
})();
