document.querySelectorAll(".add-cart").forEach((button) => {
    button.addEventListener("click", () => {
        const card = button.closest(".product-card");
        const productName = card.querySelector(".card-title").textContent.trim();
        alert(`${productName} has been added to your cart.`);
    });
});

const contactForm = document.getElementById("contactForm");
if (contactForm) {
    contactForm.addEventListener("submit", (event) => {
        event.preventDefault();
        alert("Thank you. Your message has been sent.");
        contactForm.reset();
    });
}

const feedbackForm = document.getElementById("feedbackForm");
if (feedbackForm) {
    feedbackForm.addEventListener("submit", (event) => {
        event.preventDefault();
        alert("Thank you for your feedback.");
        feedbackForm.reset();
    });
}
