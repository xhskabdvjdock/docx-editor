# محرر خصائص docx (Next.js)

    npm install
    npm run dev        # للتطوير: http://localhost:3000
    npm run build      # يولّد موقعاً ثابتاً في مجلد out/

ارفع محتويات مجلد `out/` على أي استضافة عادية (Apache / Nginx / cPanel ...).
إذا أردت تشغيله كتطبيق Node بدل الموقع الثابت، احذف سطر `output: "export"` من next.config.js ثم `npm run build && npm start`.
