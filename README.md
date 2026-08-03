# Coral Club — OKR Tree Builder

Живой сайт публикуется отсюда автоматически через Netlify (Git-деплой).

## Структура

- `index.html`, `bundle.js` — то, что реально публикуется (готовый, собранный сайт).
- `OKR_Tree_Builder.jsx` — исходный код приложения (React), редактируется именно этот файл.
- `entry.jsx` — точка входа для сборки.
- `favicon.ico`, `favicon-32.png`, `icon-192.png`, `icon-512.png`, `apple-touch-icon.png` — иконки сайта.

## Как пересобрать после правок в OKR_Tree_Builder.jsx

```bash
npm install
npm run build
```

Это перезапишет `bundle.js`. После этого закоммитьте и запушьте изменения — Netlify задеплоит новую версию сам.

## Настройки Netlify для этого репозитория

- **Build command:** `npm run build`
- **Publish directory:** `/` (корень репозитория)
