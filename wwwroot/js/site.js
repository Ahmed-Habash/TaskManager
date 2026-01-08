// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

// Scroll Position Persistence
document.addEventListener("DOMContentLoaded", function () {
    // Save scroll position before unload
    window.addEventListener("beforeunload", function () {
        sessionStorage.setItem('scrollPos', window.scrollY);
    });
});
