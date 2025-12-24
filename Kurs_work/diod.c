#include "diod.h"

Sensor sensors[MAX_SENSORS];
SensorConfig sensor_config[MAX_SENSORS];

void Init_Sensors(void) {
    for (uint8_t i = 0; i < MAX_SENSORS; ++i) {
        sensors[i].raw_temp = 0x0;
        sensors[i].temp = 0.0;
        sensors[i].crc8_rom = 0x0;
        sensors[i].crc8_data = 0x0;
        sensors[i].crc8_rom_error = 0x0;
        sensors[i].crc8_data_error = 0x0;
        
        for (uint8_t j = 0; j < 8; j++) {
            sensors[i].ROM_code[j] = 0x00;
        }
        for (uint8_t q = 0; q < 9; q++) {
            sensors[i].scratchpad_data[q] = 0x00;
        }
    }
}

void ApplySensorResolution(uint8_t sensor_index) {
    if(!sensor_config[sensor_index].changed) return;
    if(sensors[sensor_index].ROM_code[0] == 0) return;
    
    uint8_t config_byte;
    
    switch(sensor_config[sensor_index].resolution) {
        case 9:
            config_byte = RESOLUTION_9BIT;
            break;
        case 10:
						config_byte = RESOLUTION_10BIT;
            break;
				case 11:
						config_byte = RESOLUTION_11BIT;
            break;
        default:
            config_byte = RESOLUTION_12BIT;
            break;
    }
    
    ds18b20_Reset();
    ds18b20_MatchRom(sensors[sensor_index].ROM_code);
    ds18b20_WriteByte(WRITE_SCRATCHPAD);
    ds18b20_WriteByte(0x64);     
    ds18b20_WriteByte(0xFF9E);        
    ds18b20_WriteByte(config_byte); 
    
    sensor_config[sensor_index].changed = 0;
}

void diod(void) {
    uint8_t devCount = 0, i = 0;
	
    static uint8_t port_initialized = 0;
    if (!port_initialized) {
        ds18b20_PortInit();
        port_initialized = 1;
    }
		
		for(i = 0; i < MAX_SENSORS; i++) {
        if(sensor_config[i].changed) {
            ApplySensorResolution(i);
        }
    }
    
    if (ds18b20_Reset()) {
        return;
    }
		
    devCount = Search_ROM(0xF0, sensors);
    
    if (devCount == 0) {
        return;
    }
		
    for (i = 0; i < devCount; ++i) {
        if (!sensors[i].crc8_rom_error) {
            ds18b20_ConvertTemp(1, sensors[i].ROM_code);
						
						uint32_t conversion_delay;
            switch(sensor_config[i].resolution) {
                case 9:
                    conversion_delay = 94000;  // 94 ms
                    break;
                case 10:
                    conversion_delay = 188000;  // 94 ms
                    break;
								case 11:
                    conversion_delay = 375000;  // 94 ms
                    break;
                default:
                    conversion_delay = 750000; // 750 ms
                    break;
            }
            
            DelayMicro(conversion_delay);
					
            ds18b20_ReadStratchpad(1, sensors[i].scratchpad_data, sensors[i].ROM_code);
            sensors[i].crc8_data = Compute_CRC8(sensors[i].scratchpad_data, 8);
            sensors[i].crc8_data_error = (Compute_CRC8(sensors[i].scratchpad_data, 9) == 0) ? 0 : 1;
            
            if (!sensors[i].crc8_data_error) {
                sensors[i].raw_temp = ((uint16_t)sensors[i].scratchpad_data[1] << 8) | sensors[i].scratchpad_data[0];
								 uint16_t mask;
                switch(sensor_config[i].resolution) {
                    case 9:
                        mask = 0xFFF8;
                        break;
										case 10:
												 mask = 0xFFFC;
												break;
										case 11:
												mask = 0xFFFE;
												break;
                    default:
                        mask = 0xFFFF; 
                        break;
                }
                
                sensors[i].raw_temp &= mask;
								sensors[i].temp = sensors[i].raw_temp * 0.0625;
            }
        }
        DelayMicro(10000);
    }
}