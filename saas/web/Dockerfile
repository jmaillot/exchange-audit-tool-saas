FROM nginx:alpine
COPY index.html portal.css app.js config.js msal-browser.min.js /usr/share/nginx/html/
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
