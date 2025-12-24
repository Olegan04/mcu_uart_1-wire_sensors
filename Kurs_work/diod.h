#ifndef DIOD_H_
#define DIOD_H_

#include "stm32f10x.h"
#include "ds18b20.h"

#define MAX_SENSORS 2

typedef struct {
    uint8_t resolution;
    uint8_t changed; 
} SensorConfig;

void Init_Sensors(void);
void Init_Sensor_Configs(void); 
void diod(void);
void Process_Temperature_Data(void);
void ApplySensorResolution(uint8_t sensor_index);
uint8_t ReadSensorResolution(uint8_t sensor_index);
uint8_t ReadScratchpad(uint8_t sensor_index);

extern Sensor sensors[MAX_SENSORS];
extern SensorConfig sensor_config[MAX_SENSORS];

#endif /* DIOD_H_ */