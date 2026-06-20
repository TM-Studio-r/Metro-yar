// ================== تنظیمات ==================
// کلید API که خواستی گذاشته شده — همین رشته رو تغییر نده مگر لازم باشه
const DEEPSEEK_API_KEY = 'sk-ad08fd53222a40909db3ac2e696fd5a8';

// آدرس API اصلی (با همونی که فایل قبلی استفاده می‌کرد)
const DEEPSEEK_API_URL = 'https://api.deepseek.com/chat/completions';

// اگر از مرورگر مستقیم صدا زدن API با خطای CORS مواجه شدی، می‌تونی این پروکسی رو فعال کنی.
// پروکسی‌های عمومی محدودیت دارن؛ در محیط تولید بهتره سرور واسط داشته باشی.
const USE_PROXY = true;
const PROXY = 'https://api.allorigins.win/raw?url=';

function getApiUrl() {
    if (USE_PROXY) return PROXY + encodeURIComponent(DEEPSEEK_API_URL);
    return DEEPSEEK_API_URL;
}
// =============================================

// ساخت prompt مناسب برای کپشن‌ساز
function buildPrompt(topic, tone) {
    return `تو یک کپشن‌ساز هوشمند اینستاگرام هستی. خروجی باید مناسب پست‌های اینستاگرام باشه و در هر پاسخ دقیقاً ۳ کپشن کوتاه، خلاق و قابل استفاده تولید کن. هر کپشن را با خط جداگانه بنویس و از اموجی مناسب استفاده کن. 
موضوع: ${topic}
لحن: ${tone}
لطفاً هیچ کد یا قالب JSON ارسال نکن — فقط متن ساده کپشن‌ها.`;
}

async function generateCaption() {
    const topicInput = document.getElementById('topic');
    const toneInput = document.getElementById('tone');
    const btn = document.getElementById('generateBtn');

    const topic = topicInput.value.trim();
    const tone = toneInput.value;

    if (!topic) {
        alert('لطفاً موضوع عکس/ویدیو را وارد کنید.');
        topicInput.focus();
        return;
    }

    btn.disabled = true;
    btn.innerText = 'درحال نوشتن...';

    const resultBox = document.getElementById('resultBox');
    const outputSection = document.getElementById('outputSection');

    // آماده‌سازی بدنه درخواست؛ این ساختار انعطاف‌پذیره و چند حالت را پشتیبانی می‌کنه
    const prompt = buildPrompt(topic, tone);
    const body = {
        model: "deepseek-default",
        messages: [
            { role: "system", content: "You are a helpful assistant that creates Instagram captions in Persian." },
            { role: "user", content: prompt }
        ],
        max_tokens: 400,
        temperature: 0.9
    };

    try {
        const resp = await fetch(getApiUrl(), {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${DEEPSEEK_API_KEY}`
            },
            body: JSON.stringify(body)
        });

        if (!resp.ok) {
            // پیام‌های خطای کاربرپسند
            if (resp.status === 401 || resp.status === 403) {
                resultBox.innerText = 'خطا: دسترسی رد شد — لطفاً کلید API را بررسی کن.';
            } else if (resp.status === 429) {
                resultBox.innerText = 'خطا: نرخ درخواست زیاد است — بعداً امتحان کن یا از پروکسی/سرور واسط استفاده کن.';
            } else {
                const text = await resp.text();
                resultBox.innerText = `خطای سرور: ${resp.status}\n${text}`;
            }
            outputSection.style.display = 'block';
            return;
        }

        const data = await resp.json();

        // چند شکل متداول پاسخ رو بررسی می‌کنیم و مناسب‌ترین متن رو استخراج می‌کنیم
        let generated = '';

        if (data.choices && data.choices[0]) {
            // ساختار مشابه OpenAI: choices[0].message.content
            if (data.choices[0].message && data.choices[0].message.content) {
                generated = data.choices[0].message.content;
            } else if (data.choices[0].text) {
                generated = data.choices[0].text;
            }
        }

        // بعضی سرویس‌ها پاسخ رو مستقیماً در فیلد result یا output می‌فرستن
        if (!generated && data.result) generated = data.result;
        if (!generated && data.output) generated = data.output;
        if (!generated) {
            // fallback: stringify برای دیباگ
            generated = JSON.stringify(data, null, 2);
        }

        // پاک کردن فاصله‌های اضافه و نمایش
        resultBox.innerText = generated.trim();
        outputSection.style.display = 'block';

    } catch (err) {
        console.error(err);
        resultBox.innerText = 'خطا در اتصال به سرویس — احتمالاً مشکل CORS یا اینترنت. «برای رفع سریع»: یا USE_PROXY را true کن یا این فایل را از طریق یک سرور محلی (مثلاً با `npx http-server` یا `python -m http.server`) اجرا کن.';
        outputSection.style.display = 'block';
    } finally {
        btn.disabled = false;
        btn.innerText = 'بنویس برام! 🚀';
    }
}

function copyToClipboard() {
    const text = document.getElementById('resultBox').innerText;
    navigator.clipboard.writeText(text)
        .then(() => {
            const btn = document.getElementById('copyBtn');
            btn.innerText = '✅ کپی شد!';
            setTimeout(() => btn.innerText = 'کپی کن 📋', 2000);
        })
        .catch(() => alert('کپی در کلیپ‌بورد امکان‌پذیر نیست.'));
}

// اگر کاربر از onclick در HTML استفاده نکرده باشه، اینجا هندلر رو به دکمه اختصاص میدیم (ایمن)
document.addEventListener('DOMContentLoaded', () => {
    const btn = document.getElementById('generateBtn');
    if (btn) btn.addEventListener('click', (e) => {
        e.preventDefault();
        generateCaption();
    });
});