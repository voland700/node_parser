const { JSDOM } = require('jsdom');
const fs = require('fs');
const https = require('https');
const axios = require('axios');
const { promises: Fs } = require('fs')
const ExcelJS = require('exceljs');

const data = [];
// Случайная строка для уникалитзации имени файла при скачивании
function getString(length) {
    var result           = '';
    var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
}
/**
// Скачать файл по ссылке dir - в поддерикторию папки uploud, prefix - подстрака в названии файла
function getFile(url, dir='images', prefix=''){ 
    //Определяем расширение файла, формируем путь из дерикторий и название файла   
    let extension = url.slice(url.lastIndexOf('.') + 1);
    dir = `./upload/${dir}`;
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    let filePath = `${dir}/${prefix}_${getString(10)}.${extension}`;     
    //Функция для скачиванеи файла на диск
    function doRequest(url, filePath) {      
        https.get(url, function(res) {           
            res.on('data', function(data) {                            
                require('fs').createWriteStream(filePath, {flags:'wx'}).write(data);
            });
        });
    }
    //Функция для проверки неаличия скаченного файла на дикске
    async function exists (path) {  
        try {
            await Fs.access(path)
            return true
        } catch {
            return false
        }
    }
    //Скачиваем файл по ссылке  
    doRequest(url, filePath); 
    //Проверяем, если файл скачен на диск возвращаем строку - путь к файлу 
    return exists(filePath) ? filePath : false  
}
 */


function getFile(url, dir='images', prefix=''){ 
    //Определяем расширение файла, формируем путь из дерикторий и название файла   
    let extension = url.slice(url.lastIndexOf('.') + 1);
    dir = `./upload/${dir}`;
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    let filePath = `${dir}/${prefix}_${getString(10)}.${extension}`;     
    //Функция для скачиванеи файла на диск
    async function downloadFile(url, filePath) {

    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(filePath);
        
        https.get(url, (response) => {
            // Проверяем статус код ответа
            if (response.statusCode !== 200) {
                reject(new Error(`Ошибка загрузки: ${response.statusCode}`));
                return;
            }
            // Получаем общий размер файла для прогресса
            const totalSize = parseInt(response.headers['content-length'], 10);
            let downloadedSize = 0;
            
            response.pipe(file);
            file.on('finish', () => {
                file.close();
                resolve({ path: filePath, size: downloadedSize });
            });
            }).on('error', (err) => {
                fs.unlink(filePath, () => reject(err));
            });

            file.on('error', (err) => {
                fs.unlink(filePath, () => reject(err));
            });
        });
    }
    //Функция для проверки неаличия скаченного файла на дикске
    async function exists (path) {  
        try {
            await Fs.access(path)
            return true
        } catch {
            return false
        }
    }
    //Скачиваем файл по ссылке  
    downloadFile(url, filePath); 
    //Проверяем, если файл скачен на диск возвращаем строку - путь к файлу 
    return exists(filePath) ? filePath : false  
}




// Функция очистка HTML от мусора
function normalizeHtml(html){  
    function stripStyles(html) {       
        Array.from(html.querySelectorAll('*')).forEach(node => node.removeAttribute('style'));
        return html;
    }
    let p_color = html.querySelector("[style*='color:red']");
    if(p_color) p_color.remove();
    const links = html.querySelectorAll('a');  
    if(links.length > 0) links.forEach(link => link.parentNode.removeChild(link));
    html = stripStyles(html);
    html = html.innerHTML;
    html = html.replace(/\<(?!\/?(p|ul|li|br|b|table|tbody|tr|td|th|h1|h2|h3|h4|h5|h6|span)[ >])[^>]*\>/ig,"");
    html = html.replace('<p><br/></p>', '').replace('<p></p>', '');
    html = html.replace(/^\s*[\r\n]/gm, '');
    return html;
}

async function getItem(url){
    let data = {};
    const response = await axios.request({
        method: "GET",
        url: url,
        headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        }
    });           
    const dom = new JSDOM(response.data); 
    const content =  dom.window.document;   
    
    let description = '';
    let main = ''
    let more = '' 
    let category = content.querySelector('span.ty-breadcrumbs__current').textContent;
    let html = content.getElementById('content_description');
    if(html) description = normalizeHtml(html);         
    let name = content.querySelector('h1.ty-product-block-title').textContent;
    let sku = content.querySelector('span.ty-control-group__item').textContent;
    let price = content.querySelector('span.ty-price-num').textContent;
    let images = content.querySelectorAll('.cm-image-previewer');       
    if(images){
        let imagesUrls = []
        images.forEach(item =>{
            let href = item.getAttribute('href');                
            imagesUrls.push(href);
        })
        if(imagesUrls.length>0){
            let more_images = [];
            imagesUrls.forEach((url, i) => {  
                if(i === 0){                       
                    let image = getFile(url, 'main', 'main');                    
                    if(image) main = image;
                }else{                                           
                    img = getFile(url, 'more', 'more');
                    if(img) more_images.push(img); 
                }  
            })
            if(more_images.length > 0) more = more_images.join(',');
        } 
    }        
    let properties = [];
    let propertiesDom = content.getElementById('content_features');
    if(propertiesDom){
        let props = propertiesDom.querySelectorAll('.ty-product-feature');
        props.forEach( item => {
            let lable = item.querySelector('.ty-product-feature__label').textContent;
            let value = item.querySelector('.ty-product-feature__value').textContent;
            properties.push({lable: lable, value: value});
        });
    } 
    data = {url: url, category: category, name: name, sku: sku, price: price,  main: main, more: more, description: description, properties: properties}; 
    if (Object.keys(data).length > 0) {
        //console.log(JSON.stringify(data, null, ' '));
        console.log(` - ${data['name']}`)  
        return data;              
    }
}

// Чтение текстового файла с спиком сылок для парсинга построчно - данные в маасив (файл './source.txt' - в корне)
function getUrlFromFile(file_path){
    let list = [];
    let lines = fs.readFileSync(file_path, 'utf8').split('\n');
    if(lines) {
        lines.forEach(line => {
            list.push(line.trim());
        })
    }
    if(list.length > 0 ){
        return list;
    } else {
        return false;
    }
}

// Запись данных в JSON файл, принимает данные и название для файла.
function dataToJson(data, file_name){
    fs.writeFileSync(`${file_name}.json`, JSON.stringify(data, null, ' '));
}

// Получние данных из json - файла
function readJsonFile(file_path) { 
    const data = fs.readFileSync(file_path);
    return JSON.parse(data);
}


const getData = async function(urls) {  
    
    const products = await function(urls) {
        
        let products_data = [];
        urls.forEach(async (url) => {
            const item = await getItem(url,); 

             console.log(item);


            if(item) products_data.push(item);           
        });
        console.log(products_data);
        return products_data;
    }
    return products;
};

 async function getUrlList(){
    let urls = await  getUrlFromFile('./source.txt')
    if(!urls){
        console.log('Не получены URL -  адреса из списка ссылок');
        return false;
    } else {
        console.log(`Получено ${urls.length} ссылок. Приступаю к парсингу данных:`)
    }
    return urls;
} 


// Парсинг данных по списку ссылок из './source.txt' и запси данных в JSON - файл.
async function getDataToJsonFile(){
    let urls = await getUrlList();
    let data = [];

    for (const item of urls) {
        const result = await getItem(item);
        data.push(result);        
    }      
    dataToJson(data, 'test_data')
}
// TO DO - изменить, получить дату и извлечь уникальные имена характеристик товаров
function getProps(All_data){
    
    //let All_data =  readJsonFile('./test_data.json');
    let unique_nams = [];

    let row_template = {
        url: '',
        category: '',
        name: '',
        sku: '',
        price: '',
        main: '',
        more: '',
        description: '',
        json_properties: '',
    }

    let first_row = {
        url: 'Ссылка',
        category: 'Категория',
        name: 'Название',
        sku: 'SKU',
        price: 'Цена',
        main: 'Главная',
        more: 'Доплнительные',
        description: 'Описание',
        json_properties: 'JSON характеристики'
    }

    let row_properties = {} ; // Формируем шаблон для поиска соответсвия колонок с названием характеристик и значений 
    

    All_data.forEach(item => {
        props = item['properties'];       
        props.forEach(prop_item => {
            unique_nams.indexOf(prop_item['lable']) === -1 ? unique_nams.push(prop_item['lable']) : false;
        })
    }) 

    unique_nams.forEach((prop, i) => {
        let key = `prop_${i+1}`;
        row_properties[key] = prop;
        row_template[key] = '';
        first_row[key] = prop;
    });

    let columns = []; // Для worksheet.columns - необходимые данные для  записи в excel - файл
    for (const [key, value] of Object.entries(first_row)) {
        columns.push({ header: value, key: key, width: 10 })
    }
    let data = [];
    
    All_data.forEach(item => {
        let item_template = {};
        for (let key in row_template) {
            item_template[key] = row_template[key];
        }
        item_template['url'] = item['url'];
        item_template['category'] = item['category']; 
        item_template['name'] = item['name'];
        item_template['sku'] = item['sku'];
        item_template['price'] = item['price'] ? item['price'].replace('\xa0', '') : 0;
        item_template['main'] = item['main'];
        item_template['more'] = item['more'];
        item_template['description'] = item['description'];
        item_template['json_properties'] = item['properties'] ? JSON.stringify(item['properties']) : '';
        if(item['properties']){
            item['properties'].forEach(props_item => {
                for (const [key, value] of Object.entries(row_properties)) {
                    if(props_item['lable'] === value) item_template[key] = props_item['value'];
                }
            });
        }
        data.push(item_template);
     })

    const file_name = `resault_${getString(10)}..xlsx`;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("data");
    worksheet.columns = columns;
    data.forEach(row => {
        worksheet.addRow(row);
    });
    workbook.xlsx.writeFile(file_name).then(() => {
        console.log('****************');
        console.log(`Данные записаны в файл: ${file_name}`);
        console.log('****************');
    });
}

// Парсинг данных по списку ссылок из './source.txt' формируем данные для записи.
async function getExcelFile(){
    let urls = await getUrlList();
    let All_data = [];
    for (const item of urls) {
        const result = await getItem(item);
        All_data.push(result);        
    }     
    getProps(All_data);
}

getExcelFile()

