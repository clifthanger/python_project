        const API_BASE = "http://127.0.0.1:6969";
        const API_KEY = "PUB_FARHORIZON";
    
        const track = document.querySelector(".carousel-track");
        const cards = document.querySelectorAll(".card");
        const container = document.querySelector(".carousel-container");
        let index = 0;
    
        function updateCarousel() {
          track.style.transform = `translateX(-${index * 100}%)`;
          container.style.height = cards[index].offsetHeight + "px";
        }
    	
    	function autoResizeInput(input) {
      const tmp = document.createElement("span");
      tmp.style.visibility = "hidden";
      tmp.style.position = "absolute";
      tmp.style.whiteSpace = "pre";
      tmp.style.font = window.getComputedStyle(input).font;
      document.body.appendChild(tmp);
    
      const updateWidth = () => {
        tmp.textContent = input.value || input.placeholder || "";
        input.style.width = tmp.scrollWidth + 30 + "px"; // +30 biar gak kepotong
      };
    
      input.addEventListener("input", updateWidth);
      updateWidth();
    }
    
    // buat select juga
    function autoResizeSelect(select) {
      const tmp = document.createElement("span");
      tmp.style.visibility = "hidden";
      tmp.style.position = "absolute";
      tmp.style.whiteSpace = "pre";
      tmp.style.font = window.getComputedStyle(select).font;
      document.body.appendChild(tmp);
    
      const updateWidth = () => {
        tmp.textContent = select.options[select.selectedIndex].text;
        select.style.width = tmp.scrollWidth + 40 + "px"; // +40 buat padding arrow
      };
    
      select.addEventListener("change", updateWidth);
      updateWidth();
    }
    
    window.addEventListener("DOMContentLoaded", () => {
      document.querySelectorAll("input").forEach(autoResizeInput);
      document.querySelectorAll("select").forEach(autoResizeSelect);
    });
    
    
        document.querySelector(".next").addEventListener("click", () => {
          index = (index + 1) % cards.length;
          updateCarousel();
        });
    
        document.querySelector(".prev").addEventListener("click", () => {
          index = (index - 1 + cards.length) % cards.length;
          updateCarousel();
        });
    
        // === Swipe support ===
        let startX = 0;
        let endX = 0;
        container.addEventListener("touchstart", (e) => {
          startX = e.touches[0].clientX;
        });
    
        container.addEventListener("touchmove", (e) => {
          endX = e.touches[0].clientX;
        });
    
        container.addEventListener("touchend", () => {
          if (startX - endX > 50) {
            index = (index + 1) % cards.length;
            updateCarousel();
          } else if (endX - startX > 50) {
            index = (index - 1 + cards.length) % cards.length;
            updateCarousel();
          }
        });
    
        window.addEventListener("load", updateCarousel);
    	// === Keyboard Arrow Support ===
    	window.addEventListener("keydown", (e) => {
    	  if (e.key === "ArrowRight") {
    		index = (index + 1) % cards.length;
    		updateCarousel();
    	  } else if (e.key === "ArrowLeft") {
    		index = (index - 1 + cards.length) % cards.length;
    		updateCarousel();
    	  }
    	});
    
    
        // === Fetch Functions ===
    async function fetchTDK() {
      const type = document.getElementById("tdk-type").value;
      const val = document.getElementById("tdk-value").value.trim();
      const box = document.getElementById("tdk-container");
    
      if (!val) return alert("Isi nomor dulu bro!");
      box.innerHTML = "<p>Loading...</p>";
    
      const url = `${API_BASE}/tdk/${type}_${val}`;
      try {
        const res = await fetch(url, { headers: { "x-api-key": API_KEY } });
        const text = await res.text();
        if (!res.ok) throw new Error(text);
    
        let data;
        try {
          data = JSON.parse(text);
        } catch {
          box.innerHTML = `<p>${text}</p>`; // fallback kalau bukan JSON
          return;
        }
    
        // Buat format tampilan rapi (NAMA: xxx)
        let html = "";
        for (const [key, value] of Object.entries(data)) {
          html += `<p><b>${key.toUpperCase()}</b> : ${value}</p>`;
        }
    
        box.innerHTML = html || "<p>Tidak ada data ditemukan.</p>";
    	
    // Pastikan card ikut resize setelah isi berubah
    const card = box.closest('.card');
    card.style.height = "auto";
    card.style.height = card.scrollHeight + "px";
    
    // Paksa container carousel juga ikut menyesuaikan tinggi card aktif
    if (cards[index] === card) {
      container.style.height = card.scrollHeight + "px";
    }
    
    // Reflow kecil untuk jaga stabilitas layout
    container.offsetHeight; // trigger reflow
    
    // Baru aktifkan scroll lagi setelah 1 frame
    box.scrollTop = 0;
    box.style.overflowY = "hidden";
    setTimeout(() => {
      box.style.overflowY = "auto";
    }, 50);
    
      } catch (err) {
        box.innerHTML = `<p style="color:#f66;">Error: ${err}</p>`;
      }
    }
    
        async function fetchNPL() {
            const type = document.getElementById("npl-type").value;
            const ktr = document.getElementById("npl-ktrbrco").value.trim();
            const box = document.getElementById("npl-result");

            if (!ktr) return alert("Isi kode dulu!");
            box.innerHTML = "<p>Loading...</p>";

            try {
                const res = await fetch(`${API_BASE}/npl/${type}/${ktr}`, { headers: { "x-api-key": API_KEY } });
                if (!res.ok) throw new Error(await res.text());

                const blob = await res.blob();
                showLightbox(URL.createObjectURL(blob));

                box.innerHTML = "<p>Berhasil memuat gambar ✅</p>";
            } catch (err) {
                box.innerHTML = `<p style="color:#f66;">Error: ${err}</p>`;
            }
        }

        async function fetchNPLDetail() {
            const type = document.getElementById("npldetail-type").value;
            const ktr = document.getElementById("npldetail-ktrbrco").value.trim();
            const box = document.getElementById("npldetail-result");

            if (!ktr) return alert("Isi kode dulu!");
            box.innerHTML = "<p>Loading...</p>";

            try {
                const res = await fetch(`${API_BASE}/npldetail/${type}/${ktr}`, { headers: { "x-api-key": API_KEY } });
                if (!res.ok) throw new Error(await res.text());

                const blob = await res.blob();
                showLightbox(URL.createObjectURL(blob));

                box.innerHTML = "<p>Berhasil memuat gambar ✅</p>";
            } catch (err) {
                box.innerHTML = `<p style="color:#f66;">Error: ${err}</p>`;
            }
        }

        async function fetchHistory() {
            const acc = document.getElementById("hist-account").value.trim();
            const start = document.getElementById("hist-start").value.trim();
            const end = document.getElementById("hist-end").value.trim();
            const box = document.getElementById("hist-result");

            if (!acc || !start || !end) return alert("Lengkapi semua field!");

            box.innerHTML = "<p>Loading...</p>";

            try {
                const url = `${API_BASE}/history?account=${acc}&start_date=${start}&end_date=${end}`;
                const res = await fetch(url, { headers: { "x-api-key": API_KEY } });

                if (!res.ok) throw new Error(await res.text());

                const blob = await res.blob();
                const link = document.createElement("a");
                link.href = URL.createObjectURL(blob);
                link.download = `history_${acc}.xlsx`;
                link.click();

                box.innerHTML = "<p>File berhasil didownload ✅</p>";
            } catch (err) {
                box.innerHTML = `<p style="color:#f66;">Error: ${err}</p>`;
            }
        }
    
        async function pushGSheet() {
          const resBox = document.getElementById("gsheet-result");
          resBox.textContent = "Mengirim ke Google Sheet...";
          try {
            const res = await fetch(`${API_BASE}/push_gsheet`, {
              method: "POST",
              headers: { "x-api-key": API_KEY },
            });
            const text = await res.text();
            if (!res.ok) throw new Error(text);
            const data = JSON.parse(text);
            resBox.textContent = JSON.stringify(data, null, 2);
          } catch (err) {
            resBox.textContent = err;
          }
        }
		
		function showLightbox(src) {
		  const lb = document.getElementById("img-lightbox");
		  const img = document.getElementById("lightbox-img");
		  img.src = src;
		  lb.style.display = "flex";
		}

		function closeLightbox() {
		  const lb = document.getElementById("img-lightbox");
		  lb.style.display = "none";
		}

		// Tutup saat klik area luar
		document.getElementById("img-lightbox").addEventListener("click", (e) => {
		  if (e.target.id === "img-lightbox") {
			closeLightbox();
		  }
		});

		function openCalendar(targetInputId) {
			const calendar = document.getElementById("hidden-calendar");

			// clear previous value
			calendar.value = "";

			calendar.onchange = () => {
				const selected = calendar.value; // format yyyy-mm-dd
				if (!selected) return;

				const yyyymmdd = selected.replace(/-/g, ""); // convert to yyyymmdd

				document.getElementById(targetInputId).value = yyyymmdd;

				// auto-resize width biar tetap rapi
				document.getElementById(targetInputId).dispatchEvent(new Event("input"));
			};

			// buka calendar
			calendar.showPicker(); // browser modern support
		}
		
		function updateDynamicPlaceholder(selectElem, inputId) {
			const option = selectElem.options[selectElem.selectedIndex];
			const ph = option.dataset.ph || "";
			document.getElementById(inputId).placeholder = ph;
		}
