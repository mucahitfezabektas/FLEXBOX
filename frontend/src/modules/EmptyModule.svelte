<div class="grid h-full gap-4 xl:grid-cols-[0.96fr_1.04fr]">
  <div class="surface-panel rounded-md p-5 shadow-enterprise">
    <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Framework Blueprint</p>
    <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">
      Yeni modul alani
    </h2>
    <p class="mt-2 text-sm leading-6 text-enterprise-text-secondary">
      Bu alan yeni subsystem pencereleri icin sade bir referans noktasidir. Yeni modul eklendiginde ortak pencere
      yapisi, launcher kaydi ve backend veya Tauri entegrasyonu buradaki sozlesmeye gore baglanir.
    </p>
  </div>

  <div class="surface-panel-alt rounded-md p-5 shadow-enterprise">
    <p class="text-data-xs font-semibold uppercase tracking-[0.2em] text-enterprise-text-muted">Minimal Contract</p>
    <div class="mt-4 space-y-3 text-sm text-enterprise-text-secondary">
      <p>
        `frontend/src/lib/modules/definitions.ts` icinde modul kimligi, baslik ve varsayilan pencere sinirlari
        tanimlanir.
      </p>
      <p>
        `frontend/src/lib/modules/registry.ts` icinde component, durum etiketi ve pencere araclari baglanir.
      </p>
      <p>
        Svelte component yalnizca kendi isiyle ilgilenir; pencere drag, resize, snap ve taskbar davranislari ortak
        framework katmani tarafindan yonetilir.
      </p>
    </div>
  </div>
</div>
