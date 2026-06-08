local json = require("luci.jsonc")

function handle_format_request(env)
    -- BAGIAN 1: AMBIL DATA GPS OTOMATIS DARI ROUTER
    -- Kita buat fungsi lokal untuk tanya ke sistem router via 'ubus'
    local function get_gps_coordinates()
        -- Jalankan perintah sistem untuk ambil info GPS
        local handle = io.popen("ubus call gps info")
        local result = handle:read("*a")
        handle:close()
        
        -- Parse hasil JSON dari sistem GPS
        local gps_data = json.parse(result)
        
        -- Ambil lat/lon, jika belum lock satelit kasih nilai 0
        local lat = (gps_data and gps_data.latitude) or 0
        local lon = (gps_data and gps_data.longitude) or 0
        
        return lat, lon
    end

    -- Panggil fungsi di atas
    local my_lat, my_lon = get_gps_coordinates()
    
    -- BAGIAN 2: LOGIKA WAKTU (Auto Rounding 5 Menit)
    local now = os.time()

    local interval = 300 -- Sesuaikan dengan periode pengiriman (300 detik)
    
    -- Rumus Matematika: Bulatkan ke kelipatan 300 detik terdekat
    -- Logika: (Waktu + Setengah Interval) / Interval, dibulatkan ke bawah, dikali Interval lagi.
    local rounded_epoch = math.floor((now + (interval / 2)) / interval) * interval
    
    -- Format menjadi String "YYYY-MM-DD HH:MM:SS"
    local formatted_time = os.date("%Y-%m-%d %H:%M:%S", rounded_epoch)

    -- BAGIAN 3: SUSUN STRUKTUR JSON FINAL
    local output = {
        kd_hardware = "R0001", -- Sesuaikan dengan kd_hardware nya
        tlocal = formatted_time,
        latitude = my_lat,
        longitude = my_lon,
        sensors = {
            airpressure = "",
            airhumidity = "",
            airtemperature = "",
            windspeed = "",
            winddirection = ""
        }
    }

    -- Fungsi pembersih data string "[123]" -> angka 123
    local function parse_value(raw_string)
        if raw_string then
            local clean = string.gsub(raw_string, "%[", "")
            clean = string.gsub(clean, "%]", "")
            return tonumber(clean) -- Mengembalikan angka atau nil jika gagal
        end
        return nil -- Kembalikan nil jika data rusak
    end

    -- BAGIAN 4: ISI DATA SENSOR MODBUS KE DALAM "output.sensors"
    -- Loop hanya akan mengupdate nilai jika datanya benar-benar ada
    for _, item in pairs(env) do
        if item.name and item.data then
            local val = parse_value(item.data)
            
            -- Jika val berhasil didapat (bukan nil), timpa nilai default "" tadi dengan angka
            if val then
                -- Perhatikan: kita memasukkannya ke 'output.sensors', bukan langsung ke 'output'
                if item.name == "windspeed" then output.sensors["windspeed"] = val
                elseif item.name == "winddirection" then output.sensors["winddirection"] = val
                elseif item.name == "airtemperature" then output.sensors["airtemperature"] = val 
                elseif item.name == "airhumidity" then output.sensors["airhumidity"] = val
                elseif item.name == "airpressure" then output.sensors["airpressure"] = val
                end
            end
        end
    end

    return json.stringify(output)
end